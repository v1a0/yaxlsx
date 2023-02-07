from .utils import safe_str


class XlsxFormula:
    def __init__(self, func_name, *items):
        self.__func_name = func_name
        self.__val = ', '.join(safe_str(v) for v in items)
        self.__addon = ''

    def __math__(self, operator: str, other):
        self.__addon += f" {operator} {safe_str(other)}"
        return self

    def __add__(self, other):
        return self.__math__('+', other)

    def __sub__(self, other):
        return self.__math__('-', other)

    def __mul__(self, other):
        return self.__math__('*', other)

    def __truediv__(self, other):
        return self.__math__('/', other)

    def __iadd__(self, other):
        self.__add__(other)

    def __isub__(self, other):
        self.__sub__(other)

    def __imul__(self, other):
        self.__mul__(other)

    def __itruediv__(self, other):
        self.__truediv__(other)

    def __hash__(self):
        return hash((self.__func_name, self.__val, self.__addon))

    @property
    def value(self):
        return self.__val

    @property
    def func_name(self):
        return self.__func_name

    def __str__(self):
        return f"{self.func_name}({self.value}){self.__addon}"

    def eq(self):
        return f"{{={self.__str__()}}}"
