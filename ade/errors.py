class AdeError(Exception):
    """Базовая ошибка ADE (сообщение рассчитано на человека)."""


class MarkupError(AdeError):
    """Ошибка разметки библиотеки."""


class PlanError(AdeError):
    """Ошибка плана (plan.json)."""


def require_file(path, what: str = "Файл"):
    """Понятная ошибка вместо трейсбека, если файла нет или он не файл."""
    from pathlib import Path
    p = Path(path)
    if not p.is_file():
        raise AdeError(f"{what} не найден: {p}")
    return p
