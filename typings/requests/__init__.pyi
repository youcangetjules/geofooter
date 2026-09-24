from typing import Any, Mapping, MutableMapping

class Response:
    status_code: int
    text: str
    content: bytes
    def json(self) -> Any: ...

def get(
    url: str,
    params: Mapping[str, Any] | None = ...,
    headers: Mapping[str, str] | None = ...,
    timeout: float | tuple[float, float] | None = ...,
    **kwargs: Any,
) -> Response: ...

def post(
    url: str,
    data: Any = ...,
    json: Any = ...,
    headers: Mapping[str, str] | None = ...,
    timeout: float | tuple[float, float] | None = ...,
    **kwargs: Any,
) -> Response: ...
