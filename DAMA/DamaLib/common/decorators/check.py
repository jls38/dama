from dataclasses import dataclass
from functools import wraps
from inspect import signature

def check_dataclass_input(cls:dataclass):
    origin_init = cls.__init__
    @wraps(cls)
    def __init__ (*args):
        for i, key in enumerate(cls.__annotations__):
                if key not in ('self'):
                    if not any([isinstance(args[i+1], cls.__annotations__[key])]):
                        error_message = f"""From {cls.__module__}.{cls.__qualname__}, argument '{key}' types error - Expected: {cls.__annotations__[key]} """
                        raise TypeError(error_message)

        origin_init(*args)
    cls.__init__ = __init__
    return cls

def check_method_input(exclude):
    def decorate(method):
        @wraps(method)
        def check_type (*args, **kwargs):
            _ignore = ('self',)
            if type(exclude)!=tuple:
                raise TypeError(f'wrong type: {method.__module__}.{method.__qualname__}')
            
            excl = (*_ignore, *exclude)
            
            if (len(excl)!= (len(exclude)+len(_ignore))):
                raise ValueError(f'missing coma: {method.__module__}.{method.__qualname__}')
            
            sig = signature(method)
            for i, arg_name in enumerate(sig.parameters):
                if arg_name not in excl:
                    if not any([isinstance(args[i], sig.parameters[arg_name].annotation)]):
                        error_message = f"""Types error : {method.__module__}.{method.__qualname__}, argument '{arg_name}' (type: {type(arg_name)})  - Expected: {sig.parameters[arg_name].annotation} """
                        raise TypeError(error_message)
        
            res = method(*args, **kwargs)
            return res
        return check_type
    return decorate