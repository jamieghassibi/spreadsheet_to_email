max_len = 80

def inspect(func):
    """Decorator: Inspect the input and output of a function."""
    display = True
    def decorator(*args, **kwargs):
        nonlocal display
        if display:
            display = False
            print('_'*max_len)
            print(f'{func.__name__}')
            print(f'\targs   = {str(args):.{max_len}}')
            print(f'\tkwargs  = {str(kwargs):.{max_len}}')
            result = func(*args, **kwargs)
            print(f'\tresult = {str(result):.{max_len}}\n')
            print('_'*max_len)
        else:
            result = func(*args, **kwargs)
        return result
    return decorator

def apply_decorator(scope):
    """Apply inspect decorator to all functions in scope."""
    for key, val in dict(scope).items():
        if callable(val):
            scope[key] = inspect(val)
