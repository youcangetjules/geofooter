from typing import Any, Optional

class Country:
    name: str
    alpha_2: str
    alpha_3: str

class Countries:
    def get(self, **kwargs: Any) -> Optional[Country]: ...

countries: Countries
