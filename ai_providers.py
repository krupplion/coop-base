# -*- coding: utf-8 -*-
"""
ORG_NAME协同 · 多云服务商 AI 发票识别适配器
================================================
支持以下云服务商的多模态发票识别，并提供统一的调用入口：

  - volcano  : 火山引擎（Volcano Engine / 豆包 Doubao Vision，Ark 服务）
  - bailian  : 阿里百炼（Ali Bailian / 通义千问 qwen-vl，DashScope 原生多模态接口）
  - tencent  : 腾讯云（Tencent Cloud / 混元 Hunyuan Vision，TC3-HMAC-SHA256 签名原生接口）
  - openai   : 兼容 OpenAI Chat Completions 的任意服务（DeepSeek / 智谱 / OpenAI 等）

安全设计（核心目标：密钥不硬编码、不入库明文、不随包泄露）：
  1. 密钥解析优先级：环境变量  >  本地配置文件(ai_providers.json)  >  应用默认（无密钥）
  2. 任何密钥都不会写死在源码中，也不会写入数据库明文；配置文件中仅存放在
     用户可写目录（开发态 data/，发布态 %LOCALAPPDATA%\\APP_DATA_NAME\\），
     该目录不随 exe 打包分发。
  3. 对外返回给前端的配置一律脱敏，仅在运行时由本模块在内存中拼装真实密钥。

各服务商的「请求参数 / 鉴权方式 / 返回结果处理」见下方 PROVIDER_PRESETS[x]["protocol"]，
前端「接口协议说明」面板即由该字段渲染。
"""

import os
import sys
import json
import base64
import hashlib
import hmac
import datetime
import urllib.request
import urllib.error


# ---------------------------------------------------------------- 配置文件路径
# 与 app.py 的 DATA_DIR 保持一致：发布态落在 %LOCALAPPDATA%\\APP_DATA_NAME，开发态落在项目 data/。
def _config_dir():
    if getattr(sys, "frozen", False):
        d = os.environ.get("COOP_DATA")
        if not d:
            _name = os.environ.get("COOP_APP_DATA", "协同录入系统")
            d = os.path.join(os.path.expandvars("%LOCALAPPDATA%"), _name)
        return d
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")


CONFIG_PATH = os.path.join(_config_dir(), "ai_providers.json")
# 数据库设置兜底读取（仅用于非密钥项，如激活的 provider）；由 app.py 在启动时注册。
_DB_GETTER = None


def set_db_getter(func):
    """app.py 注册一个读取 settings 表的回调，仅用于非密钥项兜底。"""
    global _DB_GETTER
    _DB_GETTER = func


def _db_get(key, default=""):
    if _DB_GETTER:
        try:
            return _DB_GETTER(key, default)
        except Exception:
            return default
    return default


# ---------------------------------------------------------------- 服务商预设
# 每个预设包含：默认 endpoint / 模型、该服务商专属的环境变量前缀、是否使用 AK/SK、
# 以及 protocol 文档（请求参数 / 鉴权方式 / 返回结果处理）。
PROVIDER_PRESETS = {
    "volcano": {
        "name": "火山引擎（豆包视觉）",
        "desc": "字节跳动火山引擎方舟（ARK）平台的多模态视觉模型，兼容 OpenAI Chat Completions 协议。",
        "default_base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "default_model": "doubao-visual-pro-32k",
        "auth_kind": "bearer",          # 单密钥 Bearer
        "env_prefix": "VOLCANO",
        "protocol": {
            "endpoint": "POST {base_url}/chat/completions",
            "request_params": [
                ("model", "模型 ID，如 doubao-visual-pro-32k（在方舟控制台开通后获得）"),
                ("messages", "用户消息，content 为数组，含一个 text 文本提示 + 一个 image_url 图片（base64 data URI）"),
                ("temperature", "固定 0，保证字段抽取稳定"),
            ],
            "auth": "HTTP 头 Authorization: Bearer <ARK_API_KEY>，密钥为火山引擎方舟 API Key（形如 AKxxx...）。",
            "response": "解析 payload['choices'][0]['message']['content']，按正则抽取首个 JSON 对象，取 invoice_no / seller / amount / tax_amount / invoice_date / category。",
        },
    },
    "bailian": {
        "name": "阿里百炼（通义千问视觉）",
        "desc": "阿里云百炼平台 DashScope 原生多模态生成接口，使用 qwen-vl 系列视觉模型。",
        "default_base_url": "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation",
        "default_model": "qwen-vl-max-latest",
        "auth_kind": "bearer",
        "env_prefix": "BAILIAN",
        "protocol": {
            "endpoint": "POST https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation",
            "request_params": [
                ("model", "视觉模型名，如 qwen-vl-max-latest / qwen-vl-plus"),
                ("input.messages", "数组，role=user，content 为数组：{image: 'data:image/png;base64,...'} 与 {text: '提示词'}"),
                ("X-DashScope-Async", "可选请求头，设为 enable 走异步任务；本适配器走同步返回。"),
            ],
            "auth": "HTTP 头 Authorization: Bearer <DASHSCOPE_API_KEY>（阿里云百炼 API Key，sk- 开头）。",
            "response": "解析 payload['output']['choices'][0]['message']['content'][0]['text']，按正则抽取首个 JSON 对象。",
        },
    },
    "tencent": {
        "name": "腾讯云（混元视觉）",
        "desc": "腾讯云混元大模型 Hunyuan-Vision，使用腾讯云标准 TC3-HMAC-SHA256 签名鉴权。",
        "default_base_url": "https://hunyuan.tencentcloudapi.com",
        "default_model": "hunyuan-vision",
        "default_region": "ap-guangzhou",
        "auth_kind": "tc3",            # AK/SK + TC3 签名
        "service": "hunyuan",
        "action": "ChatCompletions",
        "version": "2023-09-01",
        "env_prefix": "TENCENT",
        "protocol": {
            "endpoint": "POST https://hunyuan.tencentcloudapi.com  (Host: hunyuan.tencentcloudapi.com)",
            "request_params": [
                ("Model", "视觉模型名，如 hunyuan-vision"),
                ("Messages", "数组，Role=user，Contents 为数组：{Type:'image_url', ImageUrl:{Url:'data:image/png;base64,...'}} 与 {Type:'text', Text:'提示词'}"),
                ("X-TC-Action", "固定 ChatCompletions"),
                ("X-TC-Version", "固定 2023-09-01"),
                ("X-TC-Timestamp", "当前 Unix 秒级时间戳（UTC）"),
                ("X-TC-Region", "地域，如 ap-guangzhou"),
            ],
            "auth": "TC3-HMAC-SHA256 签名：用 SecretKey 对『日期/服务名/tc3_request』三级派生签名密钥，"
                    "对『方法+URI+规范化头+签名头+SHA256(正文)』做 SHA256 得到待签串，最终生成 "
                    "Authorization: TC3-HMAC-SHA256 Credential=SecretId/日期/service/tc3_request, "
                    "SignedHeaders=content-type;host, Signature=...。需 SecretId + SecretKey（腾讯云 API 密钥对）。",
            "response": "解析 payload['Response']['Choices'][0]['Message']['Content']，按正则抽取首个 JSON 对象。",
        },
    },
    "openai": {
        "name": "兼容 OpenAI（自定义 / DeepSeek / 智谱 等）",
        "desc": "任意兼容 OpenAI Chat Completions 协议的多模态服务，用于自托管或第三方兼容网关。",
        "default_base_url": "https://api.openai.com/v1",
        "default_model": "gpt-4o-mini",
        "auth_kind": "bearer",
        "env_prefix": "OPENAI",
        "protocol": {
            "endpoint": "POST {base_url}/chat/completions",
            "request_params": [
                ("model", "目标模型名"),
                ("messages", "用户消息，content 为数组，含 text 提示 + image_url 图片（base64 data URI）"),
                ("temperature", "固定 0"),
            ],
            "auth": "HTTP 头 Authorization: BeBearer <API_KEY>。",
            "response": "解析 payload['choices'][0]['message']['content']，按正则抽取首个 JSON 对象。",
        },
    },
}

DEFAULT_ACTIVE = "bailian"


# ---------------------------------------------------------------- 配置加载
def _mask(secret):
    if not secret:
        return ""
    if len(secret) <= 8:
        return "****"
    return secret[:4] + "****" + secret[-4:]


def load_config_file():
    """读取本地配置文件（可能不存在）。"""
    if not os.path.exists(CONFIG_PATH):
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_config_file(data):
    """写入配置文件，确保目录存在且权限为当前用户私有。"""
    os.makedirs(_config_dir(), exist_ok=True)
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except Exception:
        pass


def _protocol_label(auth_kind):
    return {
        "bearer": "Bearer 鉴权（兼容 OpenAI Chat Completions）",
        "tc3": "TC3-HMAC-SHA256 签名（SecretId + SecretKey）",
        "dashscope": "DashScope 原生多模态网关",
    }.get(auth_kind, auth_kind)


def _format_protocol(preset):
    """将结构化 protocol 文档格式化为多行可读文本，供前端『接口协议说明』面板渲染。"""
    pr = preset.get("protocol", {})
    lines = []
    if pr.get("endpoint"):
        lines.append("接口地址：" + pr["endpoint"])
    if pr.get("request_params"):
        lines.append("请求参数：")
        for k, v in pr["request_params"]:
            lines.append("  • %s：%s" % (k, v))
    if pr.get("auth"):
        lines.append("鉴权方式：" + pr["auth"])
    if pr.get("response"):
        lines.append("返回结果处理：" + pr["response"])
    return "\n".join(lines)


def list_providers():
    """返回所有服务商预设（不含任何密钥），并为前端协议面板补充 protocol 短标签与 detail 长说明。"""
    out = {}
    for pid, p in PROVIDER_PRESETS.items():
        out[pid] = {
            "id": pid,
            "name": p["name"],
            "desc": p["desc"],
            "auth_kind": p["auth_kind"],
            "default_model": p.get("default_model", ""),
            "protocol": _protocol_label(p["auth_kind"]),
            "detail": _format_protocol(p),
        }
    return out


def get_active_provider():
    """返回当前激活的服务商 ID（配置文件 > 数据库设置 > 默认）。"""
    cfg = load_config_file()
    if cfg.get("active") in PROVIDER_PRESETS:
        return cfg["active"]
    db_active = _db_get("ai_provider", "")
    if db_active in PROVIDER_PRESETS:
        return db_active
    return DEFAULT_ACTIVE


def resolve_provider(pid):
    """解析某服务商的完整有效配置（默认值 < 配置文件 < 环境变量）。
    返回 dict：provider_id, name, base_url, model, region, auth_kind, api_key, secret_id, secret_key。
    密钥仅在此时于内存中拼装，绝不落库。"""
    preset = PROVIDER_PRESETS.get(pid)
    if not preset:
        pid = DEFAULT_ACTIVE
        preset = PROVIDER_PRESETS[pid]

    cfg = {
        "provider_id": pid,
        "name": preset["name"],
        "auth_kind": preset["auth_kind"],
        "base_url": preset.get("default_base_url", ""),
        "model": preset.get("default_model", ""),
        "region": preset.get("default_region", ""),
        "api_key": "",
        "secret_id": "",
        "secret_key": "",
    }

    # 1) 配置文件
    file_cfg = load_config_file().get("providers", {}).get(pid, {})
    for k in ("base_url", "model", "region", "api_key", "secret_id", "secret_key"):
        if file_cfg.get(k):
            cfg[k] = file_cfg[k]

    # 2) 环境变量（最高优先级，便于容器 / CI 注入）
    pfx = preset.get("env_prefix", pid.upper())
    def env(name):
        # 形如 COOP_VOLCANO_API_KEY 或 VOLCANO_API_KEY
        return os.environ.get("COOP_%s_%s" % (pfx, name)) or os.environ.get("%s_%s" % (pfx, name))
    if preset["auth_kind"] == "tc3":
        if env("SECRET_ID"):
            cfg["secret_id"] = env("SECRET_ID")
        if env("SECRET_KEY"):
            cfg["secret_key"] = env("SECRET_KEY")
    else:
        if env("API_KEY"):
            cfg["api_key"] = env("API_KEY")
    if env("BASE_URL"):
        cfg["base_url"] = env("BASE_URL")
    if env("MODEL"):
        cfg["model"] = env("MODEL")
    if env("REGION"):
        cfg["region"] = env("REGION")

    # 兜底：通用 OPENAI 风格的环境变量
    if not cfg["api_key"] and preset["auth_kind"] != "tc3":
        cfg["api_key"] = os.environ.get("COOP_AI_API_KEY") or os.environ.get("AI_API_KEY") or ""
    if not cfg["model"]:
        cfg["model"] = _db_get("ai_model", cfg["model"])
    if not cfg["base_url"] and preset["auth_kind"] != "tc3":
        db_url = _db_get("ai_base_url", "")
        if db_url:
            cfg["base_url"] = db_url

    return cfg


def public_provider_state():
    """返回给前端的安全状态：各服务商是否「已配置」、激活项、脱敏信息，绝不含明文密钥。"""
    cfg = load_config_file()
    providers = {}
    for pid in PROVIDER_PRESETS:
        rc = resolve_provider(pid)
        if rc["auth_kind"] == "tc3":
            configured = bool(rc["secret_id"] and rc["secret_key"])
        else:
            configured = bool(rc["api_key"])
        providers[pid] = {
            "id": pid,
            "name": rc["name"],
            "auth_kind": rc["auth_kind"],
            "configured": configured,
            "model": rc["model"],
            "region": rc["region"],
            "base_url": rc["base_url"],
            # 脱敏展示：仅显示前缀与长度，绝不暴露明文
            "api_key_mask": _mask(rc["api_key"]),
            "secret_id_mask": _mask(rc["secret_id"]),
        }
    return {
        "active": get_active_provider(),
        "providers": providers,
        "config_path": CONFIG_PATH,
    }


# ---------------------------------------------------------------- 适配器基类
class ProviderAdapter:
    """统一适配器接口：build_body / auth_headers / parse / call。"""

    def __init__(self, cfg):
        self.cfg = cfg

    def endpoint(self):
        raise NotImplementedError

    def auth_headers(self):
        raise NotImplementedError

    def build_body(self, image_b64, mime, prompt):
        raise NotImplementedError

    def parse(self, payload):
        raise NotImplementedError

    @staticmethod
    def _extract_json(text):
        import re
        if not text:
            return {}
        m = re.search(r"\{.*\}", text, re.S)
        if not m:
            return {}
        try:
            return json.loads(m.group())
        except Exception:
            return {}

    def call(self, image_b64, mime, prompt, timeout=60):
        url = self.endpoint()
        body = json.dumps(self.build_body(image_b64, mime, prompt)).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        headers.update(self.auth_headers())
        req = urllib.request.Request(url, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return self.parse(json.loads(resp.read()))
        except urllib.error.HTTPError as e:
            detail = ""
            try:
                detail = e.read().decode("utf-8", "ignore")[:300]
            except Exception:
                pass
            raise RuntimeError("HTTP %s: %s" % (e.code, detail))
        except Exception as e:
            raise RuntimeError("调用失败: %s" % e)


class OpenAICompatibleAdapter(ProviderAdapter):
    """兼容 OpenAI Chat Completions 的服务（火山引擎 Ark / 通用 OpenAI 风格）。"""

    def endpoint(self):
        return self.cfg["base_url"].rstrip("/") + "/chat/completions"

    def auth_headers(self):
        return {"Authorization": "Bearer " + (self.cfg.get("api_key") or "")}

    def build_body(self, image_b64, mime, prompt):
        return {
            "model": self.cfg.get("model") or "gpt-4o-mini",
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url",
                     "image_url": {"url": "data:%s;base64,%s" % (mime, image_b64)}},
                ],
            }],
            "temperature": 0,
        }

    def parse(self, payload):
        text = payload["choices"][0]["message"]["content"]
        return self._extract_json(text)


class AliBailianAdapter(ProviderAdapter):
    """阿里百炼 DashScope 原生多模态生成接口（非 OpenAI 兼容模式）。"""

    def endpoint(self):
        return self.cfg.get("base_url") or \
            "https://dashscope.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation"

    def auth_headers(self):
        return {
            "Authorization": "Bearer " + (self.cfg.get("api_key") or ""),
            "X-DashScope-Async": "disable",
        }

    def build_body(self, image_b64, mime, prompt):
        return {
            "model": self.cfg.get("model") or "qwen-vl-max-latest",
            "input": {
                "messages": [{
                    "role": "user",
                    "content": [
                        {"image": "data:%s;base64,%s" % (mime, image_b64)},
                        {"text": prompt},
                    ],
                }]
            },
        }

    def parse(self, payload):
        text = payload["output"]["choices"][0]["message"]["content"][0]["text"]
        return self._extract_json(text)


class TencentHunyuanAdapter(ProviderAdapter):
    """腾讯云混元 TC3-HMAC-SHA256 签名原生接口。"""

    def endpoint(self):
        return self.cfg.get("base_url") or "https://hunyuan.tencentcloudapi.com"

    def auth_headers(self):
        return _tc3_sign(
            secret_id=self.cfg.get("secret_id", ""),
            secret_key=self.cfg.get("secret_key", ""),
            host="hunyuan.tencentcloudapi.com",
            service=self.cfg.get("service", "hunyuan"),
            action=self.cfg.get("action", "ChatCompletions"),
            version=self.cfg.get("version", "2023-09-01"),
            region=self.cfg.get("region") or "ap-guangzhou",
            payload=json.dumps(self._last_body, ensure_ascii=False),
        )

    def build_body(self, image_b64, mime, prompt):
        body = {
            "Model": self.cfg.get("model") or "hunyuan-vision",
            "Messages": [{
                "Role": "user",
                "Contents": [
                    {"Type": "image_url",
                     "ImageUrl": {"Url": "data:%s;base64,%s" % (mime, image_b64)}},
                    {"Type": "text", "Text": prompt},
                ],
            }],
        }
        self._last_body = body
        return body

    def parse(self, payload):
        text = payload["Response"]["Choices"][0]["Message"]["Content"]
        return self._extract_json(text)


def _build_adapter(cfg):
    pid = cfg["provider_id"]
    if pid == "bailian":
        return AliBailianAdapter(cfg)
    if pid == "tencent":
        return TencentHunyuanAdapter(cfg)
    # volcano / openai 及未知兜底：OpenAI 兼容
    return OpenAICompatibleAdapter(cfg)


# ---------------------------------------------------------------- TC3 签名
def _tc3_sign(secret_id, secret_key, host, service, action, version, region, payload):
    """腾讯云 TC3-HMAC-SHA256 签名，返回 HTTP 头字典。"""
    algorithm = "TC3-HMAC-SHA256"
    ts = int(datetime.datetime.now(datetime.timezone.utc).timestamp())
    date = datetime.datetime.fromtimestamp(ts, datetime.timezone.utc).strftime("%Y-%m-%d")

    hashed_payload = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    canonical_headers = "content-type:application/json\nhost:%s\n" % host
    signed_headers = "content-type;host"
    canonical_request = "\n".join([
        "POST", "/", "", canonical_headers, signed_headers, hashed_payload
    ])

    credential_scope = "%s/%s/tc3_request" % (date, service)
    string_to_sign = "\n".join([
        algorithm, str(ts), credential_scope,
        hashlib.sha256(canonical_request.encode("utf-8")).hexdigest()
    ])

    def _hmac(key, msg):
        return hmac.new(key, msg.encode("utf-8"), hashlib.sha256).digest()

    secret_date = _hmac(("TC3" + secret_key).encode("utf-8"), date)
    secret_service = _hmac(secret_date, service)
    secret_signing = _hmac(secret_service, "tc3_request")
    signature = hmac.new(secret_signing, string_to_sign.encode("utf-8"), hashlib.sha256).hexdigest()

    authorization = (
        "%s Credential=%s/%s, SignedHeaders=%s, Signature=%s"
        % (algorithm, secret_id, credential_scope, signed_headers, signature)
    )
    return {
        "Authorization": authorization,
        "X-TC-Action": action,
        "X-TC-Version": version,
        "X-TC-Timestamp": str(ts),
        "X-TC-Region": region,
    }


# ---------------------------------------------------------------- 统一入口
def build_invoice_prompt(cats=None):
    """构造发票 AI 识别提示词（与《发票 AI 识别提示词（整合版）》规范逐条对应）。

    两大识别重点（最高优先级）：
      ① 购买方/销售方名称 = 完整企业全称（行政区划+字号+行业/组织形式），禁止简称/漏字；
      ② 项目名称 = 完整保留税收分类编码前缀（*编码*具体名称），禁止删前缀/删星号。
    另含：通用禁止行为、20 字段映射表、数据勾稽校验、严格 JSON 输出格式、失败兜底。
    category 为系统特有的费用科目归集扩展字段（规范之外，不影响其余字段）。
    """
    cats = cats or []
    cat_hint = ("，以及费用科目 category（从以下列表选最匹配的一项：%s）"
                % "/".join(cats)) if cats else "，以及费用科目 category（无候选时填 null）"

    return (
        "你是一个发票信息结构化提取引擎，负责把发票（图片、扫描件、PDF 文本）识别为严格的 JSON 字段。"
        "任何字段不得遗漏、不得臆造、不得与其他字段串位。\n"
        "\n"
        "【两大识别重点（最高优先级，缺一不可）】\n"
        "重点① 购买方名称、销售方名称必须识别为完整企业全称：\n"
        "  - 从发票「购 买 方 / 销 售 方」信息区中的「名 称：」后开始读取，"
        "到「统一社会信用代码 / 纳税人识别号：」行之前的完整内容即为全称；\n"
        "  - 禁止只输出字号/简称（如「ORG_NAME」必须补全为「湖北ORG_FULL_NAME」，「洪聚财」必须补全为"
        "「湖北洪聚财建筑工程有限公司」）；禁止漏掉「湖北」「市」「公司」「有限」「事务所」「集团」等任何组成部分；\n"
        "  - 全称形态校验：全称由[行政区划]+[字号]+[行业/组织形式]组成，若识别结果仅为 2~6 字简称，"
        "判定识别失败，须回到原文重新定位「名 称：」字段；\n"
        "  - 名称与对应税号必须同源（同属购买方区或同属销售方区），禁止把销售方读入购买方、反之亦然。\n"
        "重点② 项目名称必须完整保留税收分类编码前缀：\n"
        "  - 项目名称列的列头可能有多种写法（项目名称、货物或应税劳务、服务名称、货物或应税劳务/服务名称、"
        "应税项目等），均视为「项目名称」列；\n"
        "  - 项目名称 = 该列单元格内的完整内容，格式为 *税收分类编码名称*具体货物/服务名称，"
        "例如 *生产生活服务*法律咨询；\n"
        "  - 税收分类编码前缀（星号包裹部分，如 *生产生活服务*）是项目名称的组成部分，必须原样保留："
        "禁止删除前缀、禁止去掉星号、禁止只输出星号后面的具体名称（如只输出「法律咨询」）；\n"
        "  - 星号是税收分类编码的定界符，不是 Markdown 强调符号或无关字符；OCR 若把星号识别成其他符号"
        "（如 ×、x、空格），需按上下文还原为 *；\n"
        "  - 项目名称必须与同一行的「金额、税率、税额」对齐读取，禁止把项目名称误读到「规格型号」列，"
        "也禁止把规格型号内容写入 item_name；\n"
        "  - 若发票有多行应税项目，逐行提取，合并时用「；」分隔。\n"
        "\n"
        "【通用禁止行为（出现即判定识别失败，修正后重试）】\n"
        "- 禁止把税号当作名称输出，或把名称输出到错误字段；\n"
        "- 禁止臆造发票上不存在的字段内容，无法读取的字段填 null；\n"
        "- 禁止把金额、税率、税额等数字与文字串位。\n"
        "\n"
        "【字段映射表（字段名必须严格一致）】\n"
        "invoice_type 发票类型：票面标题（如「电子发票（增值税专用发票）」，专票/普票二选一归并）；\n"
        "invoice_code 发票代码：票面「发票代码」，无则 null；\n"
        "invoice_no 发票号码：票面「发票号码」；\n"
        "invoice_date 开票日期：票面「开票日期」，格式 YYYY-MM-DD；\n"
        "buyer_name 购买方名称：购买方区「名 称：」完整全称；\n"
        "buyer_tax_id 购买方税号：购买方区「统一社会信用代码/纳税人识别号」；\n"
        "seller_name 销售方名称：销售方区「名 称：」完整全称；\n"
        "seller_tax_id 销售方税号：销售方区「统一社会信用代码/纳税人识别号」；\n"
        "item_name 项目名称：应税行项目名称列，必须完整保留税收分类编码前缀（星号包裹部分），"
        "如 *生产生活服务*法律咨询，多行用「；」分隔；\n"
        "spec 规格型号：无则 null；\n"
        "unit 单位：无则 null；\n"
        "quantity 数量：无则 null；\n"
        "unit_price 单价：无则 null；\n"
        "amount 金额(不含税)：票面「合计」金额，保留两位小数；\n"
        "tax_rate 税率/征收率：如 6%、13%、1%、免税；\n"
        "tax_amount 税额：票面「合计」税额，保留两位小数；\n"
        "total_amount 价税合计：票面「价税合计」小写金额，保留两位小数；\n"
        "total_amount_cn 价税合计大写：票面大写（如 伍万贰仟圆整）；\n"
        "drawer 开票人：票面「开票人」；\n"
        "remark 备注：票面备注栏内容，无则 null" + cat_hint + "。\n"
        "\n"
        "【数据勾稽校验（识别后必须执行）】\n"
        "1) amount + tax_amount 必须等于 total_amount；\n"
        "2) amount × tax_rate ≈ tax_amount（四舍五入至分）；\n"
        "3) 大写金额 total_amount_cn 与小写金额 total_amount 必须一致；\n"
        "任一校验不通过，在 remark 中追加标记「【勾稽异常：请人工复核】」。\n"
        "\n"
        "【失败兜底】\n"
        "- 无法读取的字段填 null，严禁编造；\n"
        "- 购买方或销售方名称识别结果不含行政区划与组织形式（判定为非全称）时，必须停止输出并提示"
        "「名称识别不完整，请人工核对」；\n"
        "- 项目名称识别结果缺失星号包裹的前缀（如只输出「法律咨询」）时，判定识别失败，"
        "须回到原文补齐税收分类编码前缀；无法补齐则提示「项目名称不完整，请人工核对」。\n"
        "\n"
        "【输出格式（严格 JSON，禁止输出任何解释性文字）】\n"
        "{\n"
        "  \"invoice_type\": \"\",\n"
        "  \"invoice_code\": null,\n"
        "  \"invoice_no\": \"\",\n"
        "  \"invoice_date\": \"YYYY-MM-DD\",\n"
        "  \"buyer_name\": \"\",\n"
        "  \"buyer_tax_id\": \"\",\n"
        "  \"seller_name\": \"\",\n"
        "  \"seller_tax_id\": \"\",\n"
        "  \"item_name\": \"\",\n"
        "  \"spec\": null,\n"
        "  \"unit\": null,\n"
        "  \"quantity\": null,\n"
        "  \"unit_price\": null,\n"
        "  \"amount\": 0.00,\n"
        "  \"tax_rate\": \"\",\n"
        "  \"tax_amount\": 0.00,\n"
        "  \"total_amount\": 0.00,\n"
        "  \"total_amount_cn\": \"\",\n"
        "  \"drawer\": \"\",\n"
        "  \"remark\": null,\n"
        "  \"category\": null\n"
        "}\n"
        "只输出上述 JSON，不要输出其他任何内容。"
    )


def recognize(provider_id=None, image_bytes=None, cats=None, timeout=60, mime="image/png"):
    """统一识别入口。

    参数：
      provider_id : 服务商 ID；为空时取当前激活项。
      image_bytes : 发票图片二进制（PNG/JPEG）。
      cats        : 费用科目列表（用于提示词中的科目归集）。
      mime        : 图片 MIME 类型（image/png / image/jpeg）。
    返回：
      dict: {invoice_type, invoice_code, invoice_no, invoice_date, buyer_name,
             buyer_tax_id, seller_name, seller_tax_id, item_name, spec, unit,
             quantity, unit_price, amount, tax_rate, tax_amount, total_amount,
             total_amount_cn, drawer, remark, category, raw}
    """
    pid = provider_id or get_active_provider()
    cfg = resolve_provider(pid)
    if cfg["auth_kind"] == "tc3":
        if not (cfg["secret_id"] and cfg["secret_key"]):
            raise RuntimeError("腾讯云密钥未配置（需 SecretId + SecretKey，可经环境变量或配置文件注入）")
    else:
        if not cfg["api_key"]:
            raise RuntimeError("API 密钥未配置（可经环境变量或配置文件 %s 注入）" % CONFIG_PATH)

    prompt = build_invoice_prompt(cats)

    image_b64 = base64.b64encode(image_bytes).decode("utf-8")
    adapter = _build_adapter(cfg)
    data = adapter.call(image_b64, mime, prompt, timeout=timeout)
    data["raw"] = json.dumps(data, ensure_ascii=False)[:2000]
    return data


def save_provider_settings(active, providers):
    """持久化服务商配置到本地配置文件（密钥仅存于此，不入库明文）。
    providers: {pid: {base_url, model, region, api_key, secret_id, secret_key}}
    带 '****' 的字段表示用户未修改，保留原值。"""
    cur = load_config_file()
    cur_providers = cur.get("providers", {})
    new_providers = {}
    for pid in PROVIDER_PRESETS:
        if pid not in providers:
            # 保留未传的服务商原有配置
            if pid in cur_providers:
                new_providers[pid] = cur_providers[pid]
            continue
        src = providers[pid] or {}
        merged = dict(cur_providers.get(pid, {}))
        for k in ("base_url", "model", "region"):
            if src.get(k) is not None and "****" not in str(src.get(k, "")):
                merged[k] = src[k]
        # 密钥：仅在用户填写且非脱敏占位时更新
        if cur_providers.get(pid, {}).get("auth_kind") == "tc3" or \
                PROVIDER_PRESETS[pid]["auth_kind"] == "tc3":
            if src.get("secret_id") and "****" not in str(src.get("secret_id", "")):
                merged["secret_id"] = src["secret_id"]
            if src.get("secret_key") and "****" not in str(src.get("secret_key", "")):
                merged["secret_key"] = src["secret_key"]
        else:
            if src.get("api_key") and "****" not in str(src.get("api_key", "")):
                merged["api_key"] = src["api_key"]
        new_providers[pid] = merged

    save_config_file({"active": active if active in PROVIDER_PRESETS else DEFAULT_ACTIVE,
                      "providers": new_providers})
    return public_provider_state()


# ---------------------------------------------------------------- 自检（可选）
def self_test():
    """结构与签名自检，不发起真实网络请求。"""
    assert get_active_provider() in PROVIDER_PRESETS
    for pid in PROVIDER_PRESETS:
        rc = resolve_provider(pid)
        assert "base_url" in rc and "auth_kind" in rc
    # 腾讯云签名结构检查（使用假密钥，仅验证头格式）
    hdrs = _tc3_sign("AKIDtest", "secretkey", "hunyuan.tencentcloudapi.com",
                    "hunyuan", "ChatCompletions", "2023-09-01", "ap-guangzhou",
                    '{"Model":"hunyuan-vision"}')
    assert hdrs["Authorization"].startswith("TC3-HMAC-SHA256 Credential=AKIDtest/")
    assert "X-TC-Timestamp" in hdrs and "X-TC-Action" in hdrs
    # 阿里构建体检查
    ba = AliBailianAdapter({"base_url": "", "model": "", "api_key": "x"})
    body = ba.build_body("BASE64DATA", "image/png", "PROMPT")
    assert body["input"]["messages"][0]["content"][0]["image"].startswith("data:image/png;base64,")
    # 火山/通用构建体检查
    oa = OpenAICompatibleAdapter({"base_url": "https://x/v1", "model": "m", "api_key": "k"})
    ob = oa.build_body("BASE64DATA", "image/jpeg", "PROMPT")
    assert ob["messages"][0]["content"][1]["image_url"]["url"].startswith("data:image/jpeg;base64,")
    return True


if __name__ == "__main__":
    print("self_test:", self_test())
    print("config_path:", CONFIG_PATH)
    print("active:", get_active_provider())
    import pprint
    pprint.pprint(public_provider_state())
