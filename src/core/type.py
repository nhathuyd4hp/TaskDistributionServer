class API:
    url: str

class UserCancelledError(Exception):
    def __init__(self, message="Task cancelled."):
        self.message = message
        super().__init__(self.message)