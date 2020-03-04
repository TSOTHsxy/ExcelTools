# -*- coding: utf-8 -*-

from functools import wraps
from json import load, dumps
from os.path import exists


def deploy(path, keys):
    """
    A decorator to register the view function for loading configuration.
    Using this decorator can simplify io operations.
    """

    if not exists(path): raise FileExistsError(
        'Configuration file {} does not exist.'.format(path)
    )

    def decorator(func):
        @wraps(func)
        def proxy(self):
            with open(path, 'r+') as file:
                # Load configuration items from a file.
                profile = load(file)

                if not set(profile.keys()).issubset(set(keys)): raise ValueError(
                    'Required configuration items are missing.'
                )

                returns = func(self, *[profile[key] for key in keys])

                # Update the specified configuration item
                # with the return value of the decorated function.
                if returns and len(keys) == len(returns):
                    for key, value in zip(keys, returns): profile[key] = value
                    file.seek(0, 0)
                    file.truncate()
                    file.write(dumps(profile, ensure_ascii=False))

        return proxy

    return decorator
