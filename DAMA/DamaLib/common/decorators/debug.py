import logging 
from functools import wraps

def debug_method (method):
    @wraps(method)
    def decorated (*args, **kwargs):
        log = logging.getLogger(method.__module__)
        res = method(*args, **kwargs)
        log.debug(f'{method.__name__}: done')
        return res
    return decorated

def debug_class(exclude):
    def decorateclass(cls):
        [setattr(cls, attr, debug_method(getattr(cls, attr))) for attr in cls.__dict__ if attr not in exclude]
        '''
        for attr in cls.__dict__:
            if attr not in exclude:
                setattr(cls, attr, DebugMethod(getattr(cls, attr)))
        '''
        return cls        
    return decorateclass