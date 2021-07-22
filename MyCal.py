def get_p1_adder(_out_ref, _input_ref):
    _key_adder = 0
    compare = _out_ref + 1 - _input_ref

    if compare > 0.5 + 1e-9:  # over 0.5
        _key_adder = 1  # 1
    elif compare > 0.25 + 1e-9:  # 0.25 ~ 0.5
        _key_adder = 0.5  # 0.2
    elif compare > 0.1 + 1e-9:  # 0.1 ~ 0.25
        _key_adder = 0.1  # 0.1
    elif compare > 0:  # 0 ~ 0.01
        _key_adder = 0  # 0
    elif compare == 0:  # 0
        _key_adder = 0  # 0
    elif compare > -0.1 + 1e-9:  # -0.1 ~ -0.01
        _key_adder = -0.1  # -0.1
    elif compare > -0.25 + 1e-9:  # -0.25 ~ -0.1
        _key_adder = -0.2
    elif compare > -0.5 + 1e-9:  # -0.5 ~ -0.25
        _key_adder = -0.5
    else:  # ~ -0.5
        _key_adder = -1
    return _key_adder


def get_input_adder(_output_now, _output_goal):
    compare = _output_goal - _output_now
    if compare > 5 + 1e-9:  # 5 ~
        adder = 1
    elif compare > 1 + 1e-9:  # 1 ~ 5
        adder = 0.5
    elif compare > 0.5 + 1e-9:  # 0.5 ~ 1
        adder = 0.2
    elif compare > 0.01 + 1e-9:  # 0.01 ~ 0.5
        adder = 0.1
    elif compare > 0 + 1e-9:  # 0 ~ 0.01
        adder = 0
    elif compare == 0:  # 0
        adder = 0
    elif compare > -0.01 + 1e-9:  # -0.01 ~ 0
        adder = 0
    elif compare > -0.5 + 1e-9:  # -0.5 ~ -0.01
        adder = -0.1
    elif compare > -1 + 1e-9:  # -1 ~ -0.5
        adder = -0.2
    elif compare > -5 + 1e-9:  # -5 ~ -1
        adder = -0.5
    else:  # ~ -5
        adder = -1
    return adder
