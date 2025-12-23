import streamlit as st


class PrefixedState:
    """Redirects all gets/sets to st.session_state under a given prefix."""

    def __init__(self, prefix: str):
        object.__setattr__(self, "_prefix", prefix)
        object.__setattr__(self, "_ss", st.session_state)

    def _k(self, name: str) -> str:
        return f"{self._prefix}:{name}"

    def __contains__(self, name: str) -> bool:
        return self._k(name) in self._ss

    def __getitem__(self, name: str):
        return self._ss[self._k(name)]

    def __setitem__(self, name: str, value):
        self._ss[self._k(name)] = value

    def get(self, name: str, default=None):
        return self._ss.get(self._k(name), default)

    def __getattr__(self, name: str):
        key = self._k(name)
        if key in self._ss:
            return self._ss[key]
        raise AttributeError(f"{name} not found in session_state[{self._prefix}]")

    def __setattr__(self, name: str, value):
        self._ss[self._k(name)] = value

    def pop(self, name: str, default=None):
        return self._ss.pop(self._k(name), default)

    def clear_namespace(self):
        to_del = [k for k in list(self._ss.keys()) if k.startswith(f"{self._prefix}:")]
        for k in to_del:
            del self._ss[k]
