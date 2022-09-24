import logging 
from functools import wraps

def DebugMethod (method):
    @wraps(method)
    def decorated (*args, **kwargs):
        log = logging.getLogger(method.__module__)
        res = method(*args, **kwargs)
        log.debug(f'{method.__name__}: done')
        return res
    return decorated

def DebugClass(exclude):
    def decorateclass(cls):
        for attr in cls.__dict__:
            if attr not in exclude:
                setattr(cls, attr, DebugMethod(getattr(cls, attr)))
        return cls        
    return decorateclass