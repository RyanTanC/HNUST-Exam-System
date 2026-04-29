"""主题数据管理（纯数据，不依赖GUI框架）."""


class Theme:
    _LIGHT = {
        "PRIMARY": "#0078d7", "PRIMARY_HOVER": "#005fa3",
        "ACCENT": "#ff9900", "BG": "#f0f0f0",
        "WHITE": "#ffffff", "TEXT": "#333333", "HINT_BG": "#e6f2ff",
        "HINT_TEXT": "#0066cc",
        "DANGER": "#ff6600", "SUCCESS": "#28a745", "MUTED": "#999999",
        "BORDER": "#e0e0e0", "NAV_ACTIVE": "#cce0ff",
        "NAV_CURRENT": "#0050a0", "NAV_ANSWERED_BG": "#d4edda",
        "NAV_ANSWERED_FG": "#155724", "NAV_MARKED_BG": "#f8d7da",
        "NAV_MARKED_FG": "#a71d2a", "NAV_HEADER_BG": "#e8f0fe",
        "SURFACE": "#f8f9fa", "CARD_BG": "#ffffff",
        "INPUT_BG": "#ffffff", "PROGRESS_BG": "#e0e0e0",
        "KB_HINT_BG": "#f8f8f8",
        "TOOLTIP_BG": "#ffffe0", "TOOLTIP_FG": "#333333",
        "NOTES_BG": "#fafbfc", "ANSWER_BG": "#e6ffe6",
        "WARN_BG": "#fff3cd", "WARN_TEXT": "#856404",
        "WARN_BORDER": "#ffc107", "HEADER_SUB_TEXT": "#bfdbfe",
        "CRASH_BG": "#ffffff", "CRASH_HEADER": "#dc2626",
    }

    _DARK = {
        "PRIMARY": "#4da6ff", "PRIMARY_HOVER": "#3d8adf",
        "ACCENT": "#ffaa33", "BG": "#1e1e1e",
        "WHITE": "#2d2d2d", "TEXT": "#e0e0e0", "HINT_BG": "#1a2a3a",
        "HINT_TEXT": "#66aaff",
        "DANGER": "#ff6644", "SUCCESS": "#44cc66", "MUTED": "#888888",
        "BORDER": "#444444", "NAV_ACTIVE": "#2a3a5a",
        "NAV_CURRENT": "#cc7700", "NAV_ANSWERED_BG": "#1a3a2a",
        "NAV_ANSWERED_FG": "#44cc66", "NAV_MARKED_BG": "#3a1a1a",
        "NAV_MARKED_FG": "#ff6644", "NAV_HEADER_BG": "#1a2a3a",
        "SURFACE": "#2a2a2a", "CARD_BG": "#2d2d2d",
        "INPUT_BG": "#333333", "PROGRESS_BG": "#333333",
        "KB_HINT_BG": "#252525",
        "TOOLTIP_BG": "#3a3a2a", "TOOLTIP_FG": "#e0e0e0",
        "NOTES_BG": "#252525", "ANSWER_BG": "#1a3a2a",
        "WARN_BG": "#3a3020", "WARN_TEXT": "#ffcc66",
        "WARN_BORDER": "#886600", "HEADER_SUB_TEXT": "#d0e8ff",
        "CRASH_BG": "#1e1e1e", "CRASH_HEADER": "#b72626",
    }

    _is_dark: bool = False
    _font_scale: float = 1.0

    # 颜色属性（通过 set_dark_mode 动态设置）
    PRIMARY = "#0078d7"
    PRIMARY_HOVER = "#005fa3"
    ACCENT = "#ff9900"
    BG = "#f0f0f0"
    WHITE = "#ffffff"
    TEXT = "#333333"
    HINT_BG = "#e6f2ff"
    HINT_TEXT = "#0066cc"
    DANGER = "#ff6600"
    SUCCESS = "#28a745"
    MUTED = "#999999"
    BORDER = "#e0e0e0"
    NAV_ACTIVE = "#cce0ff"
    NAV_CURRENT = "#0050a0"
    NAV_ANSWERED_BG = "#d4edda"
    NAV_ANSWERED_FG = "#155724"
    NAV_MARKED_BG = "#f8d7da"
    NAV_MARKED_FG = "#a71d2a"
    NAV_HEADER_BG = "#e8f0fe"
    SURFACE = "#f8f9fa"
    CARD_BG = "#ffffff"
    INPUT_BG = "#ffffff"
    PROGRESS_BG = "#e0e0e0"
    KB_HINT_BG = "#f8f8f8"
    TOOLTIP_BG = "#ffffe0"
    TOOLTIP_FG = "#333333"
    NOTES_BG = "#fafbfc"
    ANSWER_BG = "#e6ffe6"
    WARN_BG = "#fff3cd"
    WARN_TEXT = "#856404"
    WARN_BORDER = "#ffc107"
    HEADER_SUB_TEXT = "#bfdbfe"
    CRASH_BG = "#ffffff"
    CRASH_HEADER = "#dc2626"

    # 字体
    FONT = ("微软雅黑", 11)
    FONT_BOLD = ("微软雅黑", 11, "bold")
    FONT_TITLE = ("微软雅黑", 12, "bold")
    FONT_HUGE = ("微软雅黑", 16, "bold")
    FONT_SMALL = ("微软雅黑", 9)
    FONT_TINY = ("微软雅黑", 7)

    @classmethod
    def set_dark_mode(cls, enabled: bool) -> None:
        """切换深色/浅色模式."""
        cls._is_dark = enabled
        colors = cls._DARK if enabled else cls._LIGHT
        for k, v in colors.items():
            setattr(cls, k, v)

    @classmethod
    def update_fonts(cls) -> None:
        """根据 _font_scale 更新所有字体大小."""
        s = cls._font_scale
        cls.FONT = ("微软雅黑", max(8, int(11 * s)))
        cls.FONT_BOLD = ("微软雅黑", max(8, int(11 * s)), "bold")
        cls.FONT_TITLE = ("微软雅黑", max(9, int(12 * s)), "bold")
        cls.FONT_HUGE = ("微软雅黑", max(12, int(16 * s)), "bold")
        cls.FONT_SMALL = ("微软雅黑", max(7, int(9 * s)))
        cls.FONT_TINY = ("微软雅黑", max(6, int(7 * s)))

    @classmethod
    def get_current_colors(cls) -> dict:
        """获取当前主题的所有颜色."""
        return cls._DARK if cls._is_dark else cls._LIGHT
