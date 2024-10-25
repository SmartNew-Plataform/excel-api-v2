def protect(wb, ws, protect):
    password = 'password#'
    options = None

    if 'password' in protect:
        password = protect['password']
    if 'options' in protect:
        options = protect['options']
    
    ws.protect(password,options)
