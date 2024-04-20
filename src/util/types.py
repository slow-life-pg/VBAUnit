class Config:
    scenario: str

    def parse(self, jsondict: dict) -> None:
        self.scenario = jsondict["scenario"]
