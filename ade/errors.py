class AdeError(Exception):
    """Базовая ошибка ADE (сообщение рассчитано на человека)."""


class MarkupError(AdeError):
    """Ошибка разметки библиотеки."""


class PlanError(AdeError):
    """Ошибка плана (plan.json)."""
