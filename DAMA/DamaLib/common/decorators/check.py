from dataclasses import dataclass
from functools import wraps
from inspect import signature

def check_dataclass_input(cls:dataclass):
    try:
        origin_init = cls.__post_init__
    except: 
        raise Exception(f'{cls.__module__}.{cls.__qualname__} : __post_init__ method must be defined')
    @wraps(cls)
    def __post_init__ (*args):
        for key in cls.__annotations__:
            if not any([isinstance(getattr(*args, key), cls.__annotations__[key])]):
                error_message = f"Types error : {cls.__module__}.{cls.__qualname__}, {key=} (type: {type(key)}) - Expected: {cls.__annotations__[key]}"
                raise TypeError(error_message)
        origin_init(*args)
    cls.__post_init__ = __post_init__
    return cls


def check_method_input(exclude):
    def decorate(method):
        @wraps(method)
        def check_type (*args, **kwargs):
            if type(exclude) == tuple:
                excl = ('self', *exclude)
            elif type(exclude) == str:
                excl = ('self', exclude)
            else:
                raise TypeError(f'wrong type: {method.__module__}.{method.__qualname__}')

            sig = signature(method)
            for index, arg_name in enumerate(sig.parameters):
                if not arg_name in excl:
                    if not any([isinstance(args[index], sig.parameters[arg_name].annotation)]):
                        error_message = f"Types error : {method.__module__}.{method.__qualname__}, {arg_name=} (type: {type(arg_name)}) - Expected: {sig.parameters[arg_name].annotation}"
                        raise TypeError(error_message)
        
            return method(*args, **kwargs)
        return check_type
    return decorate