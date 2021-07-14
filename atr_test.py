out_ref = 0
input_ref = 0
key = -20


def get_key_adder(_out_ref, _input_ref):
    _key_adder = 0
    compare = _out_ref + 1 - _input_ref
    if compare >= 0.5:  # over 0.5
        _key_adder = 1  # 1
    elif compare >= 0.25:  # 0.25 ~ 0.5       4
        _key_adder = 0.25  # 0.25
    elif compare >= 0.1:  # 0.1 ~ 0.25        3
        _key_adder = 0.1  # 0.1
    elif compare >= 0.01:  # 0.01 ~ 0.1       2
        _key_adder = 0.05  # 0.05
    elif compare > 0:  # 0 ~ 0.01
        _key_adder = 0.01  # 0.01             1
    elif compare == 0:  # 0
        _key_adder = 0  # 0
    elif compare >= -0.01:  # -0.01 ~ 0       1
        _key_adder = -0.01  # -0.01
    elif compare >= -0.1:  # -0.1 ~ -0.01     2
        _key_adder = -0.05  # -0.05
    elif compare >= -0.25:  # -0.25 ~ -0.1    3
        _key_adder = -0.1
    elif compare >= -0.5:  # -0.5 ~ -0.25     4
        _key_adder = -0.25
    else:  # ~ -0.5
        _key_adder = -1
    return _key_adder


if __name__ == "__main__":
    print(get_key_adder(_out_ref=8.4, _input_ref=10))
