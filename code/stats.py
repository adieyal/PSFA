import math
import itertools
import functools

def percentile(N, percent, key=lambda x:x):
    """
    Find the percentile of a list of values.

    @parameter N - is a list of values. Note N MUST BE already sorted.
    @parameter percent - a float value from 0.0 to 1.0.
    @parameter key - optional key function to compute value from each element of N.

    @return - the percentile of the values
    """
    if not N:
        return None
    k = (len(N)-1) * percent
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return key(N[int(k)])
    d0 = key(N[int(f)]) * (c-k)
    d1 = key(N[int(c)]) * (k-f)
    return d0+d1

# median is 50th percentile.
median = functools.partial(percentile, percent=0.5)

def order(x):
    """
    returns the order of each element in x as a list.
    """
    L = len(x)
    rangeL = range(L)
    z = itertools.izip(x, rangeL)
    z = itertools.izip(z, rangeL) #avoid problems with duplicates.
    D = sorted(z)
    return [d[1] for d in D]

def rank(x):
    """
    Returns the rankings of elements in x as a list.
    """
    L = len(x)
    ordering = order(x)
    ranks = [0] * len(x)
    for i in range(L):
        ranks[ordering[i]] = i
    return ranks

def mean(vals):
    real_vals = [x for x in vals if x != None]
    # TODO - perhaps this isn't the best assumption?
    if len(real_vals) == 0: return 0
    return float(sum(real_vals)) / len(real_vals)
