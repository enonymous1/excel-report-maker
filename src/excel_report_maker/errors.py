"""Errors raised when a template or data payload is unusable."""


class TemplateError(ValueError):
    """A template is missing, malformed, or asks for something unsupported."""


class DataError(ValueError):
    """A data payload does not match the shape its template requires."""
