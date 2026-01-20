from typing import List

from sqlmodel import Session, select

from src.model import Error


class ErrorService:
    def __init__(self, session: Session):
        self.session = session

    def findByRunID(self, id: str) -> Error | None:
        return self.session.exec(select(Error).where(Error.run_id == id)).first()

    def findMany(self) -> List[Error]:
        return self.session.exec(select(Error)).all()
