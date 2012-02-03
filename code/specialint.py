class SpecialInt(object):
    def __init__(self, val):
        self.value = val

    def __add__(self, other):
        if other == None and self.value == None:
            return 0
        elif other == None:
            return self.value
        elif self.value == None:
            return other
        return self.value + other

    def __radd__(self, other):
        return self.__add__(other)

    def __sub__(self, other):
        if other == None or self.value == None:
            return self
        else:
            return self.value - other

    def __rsub__(self, other):
        if other == None or self.value == None:
            return self
        else:
            return other - self.value

    def __mul__(self, other):
        if other == None or self.value == None:
            return 0
        return self.value * other

    def __rmul__(self, other):
        return self.__mul__(other)

    def __div__(self, other):
        if self.value == None:
            return 0
        return self.value / other

    def __rdiv__(self, other):
        return other / self.value

    def __nonzero__(self):
        if self.value == None:
            return False
        return True

    def __gt__(self, other):
        return self.value > other

    def __ge__(self, other):
        return self.value >= other

    def __lt__(self, other):
        return self.value < other

    def __le__(self, other):
        return self.value <= other

    def __eq__(self, other):
        return self.value == other

    def __ne__(self, other):
        return self.value != other

    def __repr__(self):
        return "%s" % self.value
