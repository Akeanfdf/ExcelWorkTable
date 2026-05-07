"""Minimal node stand-in for calling register_runner functions without NodeGraphQt."""


class PropertyBagNode:
    __slots__ = ("_props",)

    def __init__(self, props: dict | None = None):
        self._props = dict(props or {})

    def get_property(self, name: str):
        return self._props.get(name)

    def set_property(self, name: str, value) -> None:
        self._props[name] = value
