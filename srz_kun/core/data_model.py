import uuid

class Person:
    def __init__(self, name: str, relationship_to_deceased: str,
                 date_of_birth: str, permanent_domicile: str, address: str,
                 is_alive: bool, waived_inheritance: bool,
                 date_of_death: str = None, id: str = None): # id is now optional and last
        self.id: str = id if id else str(uuid.uuid4())
        self.name: str = name
        self.relationship_to_deceased: str = relationship_to_deceased
        self.date_of_birth: str = date_of_birth  # ISO format YYYY-MM-DD
        self.date_of_death: str = date_of_death  # ISO format YYYY-MM-DD, or None
        self.permanent_domicile: str = permanent_domicile
        self.address: str = address
        self.is_alive: bool = is_alive
        self.waived_inheritance: bool = waived_inheritance

    def __repr__(self):
        return (f"Person(id='{self.id}', name='{self.name}', "
                f"is_alive={self.is_alive}, dob='{self.date_of_birth}')")

    def __str__(self):
        return self.name


class Relationship:
    def __init__(self, from_person_id: str, to_person_id: str, type: str):
        self.from_person_id: str = from_person_id
        self.to_person_id: str = to_person_id
        self.type: str = type  # e.g., "SPOUSE", "PARENT", "CHILD"

    def __repr__(self):
        return (f"Relationship(from='{self.from_person_id}', "
                f"to='{self.to_person_id}', type='{self.type}')")
