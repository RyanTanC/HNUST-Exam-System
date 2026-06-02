"""版本号解析与比较工具.

版本格式：yyyy.mm.dd.tttt（如 2026.05.21.1430）
- yyyy: 年
- mm:   月
- dd:   日
- tttt: 时间（如 1430 = 14:30，可选，旧格式 yyyy.mm.dd 兼容）

比较规则：先比日期（年月日），再比时间。
"""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def parse_version(version: str) -> tuple[int, ...]:
    """将版本字符串解析为可比较的整数元组.

    解析失败的段视为 0，完全无法解析的版本返回 (0, 0, 0, 0)。

    Returns
    -------
    tuple[int, int, int, int]
        (year, month, day, time)
    """
    if not version or not version.strip():
        return (0, 0, 0, 0)

    parts = version.strip().split(".")

    # 确保至少有4个元素，不足补0
    while len(parts) < 4:
        parts.append("0")

    result = []
    for p in parts[:4]:
        try:
            result.append(int(p))
        except (ValueError, TypeError):
            logger.warning("版本号段无法解析为整数: '%s' in '%s'", p, version)
            result.append(0)

    return tuple(result)


def compare_versions(local: str, remote: str) -> int:
    """比较两个版本号.

    Parameters
    ----------
    local : str
        本地版本号。
    remote : str
        远程（Gitee）版本号。

    Returns
    -------
    int
        -1: 本地比远程旧，需要更新
         0: 版本相同，无需更新
         1: 本地比远程新，无需更新
    """
    local_tuple = parse_version(local)
    remote_tuple = parse_version(remote)

    if local_tuple < remote_tuple:
        return -1
    elif local_tuple > remote_tuple:
        return 1
    return 0
