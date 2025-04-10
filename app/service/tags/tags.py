import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def add_tags(wb, ws, tags):
    logger.info(f"Chamando add_tags com tags: {tags}")
    if not isinstance(tags, dict):
        raise ValueError("'tags' deve ser um dicionário. Exemplo: {'set_row': [...], 'set_column': [...]}")
    allowed_keys = ['set_row', 'set_column', 'protect']
    invalid_keys = [key for key in tags if key not in allowed_keys]
    if invalid_keys:
        raise ValueError(f"Chaves inválidas em 'tags': {invalid_keys}. Chaves permitidas: {allowed_keys}")
    
    try:
        if 'set_row' in tags:
            logger.info(f"Aplicando set_row: {tags['set_row']}")
            set_row(wb, ws, tags['set_row'])  # Corrige para set_row (singular)
        if 'set_column' in tags:
            logger.info(f"Aplicando set_column: {tags['set_column']}")
            set_column(wb, ws, tags['set_column'])  # Corrige para set_column (singular)
        if 'protect' in tags:
            logger.info(f"Aplicando protect: {tags['protect']}")
            protect(wb, ws, tags['protect'])
    except Exception as e:
        logger.error(f"Erro ao aplicar tags: {str(e)}", exc_info=True)
        raise ValueError(f"Erro ao aplicar tags: {str(e)}")
    logger.info("add_tags concluído")

def set_row(wb, ws, rows):
    if not isinstance(rows, list):
        raise ValueError("'set_row' deve ser uma lista de dicionários. Exemplo: [{'row': 0, 'height': 30}]")
    for row_config in rows:
        if not isinstance(row_config, dict):
            raise ValueError(f"Cada item em 'set_row' deve ser um dicionário. Item inválido: {row_config}")
        if 'row' not in row_config or 'height' not in row_config:
            raise ValueError(f"Cada item em 'set_row' deve ter 'row' e 'height'. Item inválido: {row_config}")
        row = row_config['row']
        height = row_config['height']
        if not isinstance(row, int) or row < 0:
            raise ValueError(f"'row' deve ser um inteiro não negativo. Valor: {row}")
        if not isinstance(height, (int, float)) or height <= 0:
            raise ValueError(f"'height' deve ser um número positivo. Valor: {height}")
        logger.info(f"Definindo altura da linha {row} para {height}")
        ws.set_row(row, height)

def set_column(wb, ws, columns):
    if not isinstance(columns, list):
        raise ValueError("'set_column' deve ser uma lista de dicionários. Exemplo: [{'column': 0, 'width': 10}]")
    for col_config in columns:
        if not isinstance(col_config, dict):
            raise ValueError(f"Cada item em 'set_column' deve ser um dicionário. Item inválido: {col_config}")
        if 'column' not in col_config or 'width' not in col_config:
            raise ValueError(f"Cada item em 'set_column' deve ter 'column' e 'width'. Item inválido: {col_config}")
        column = col_config['column']
        width = col_config['width']
        if not isinstance(column, int) or column < 0:
            raise ValueError(f"'column' deve ser um inteiro não negativo. Valor: {column}")
        if not isinstance(width, (int, float)) or width <= 0:
            raise ValueError(f"'width' deve ser um número positivo. Valor: {width}")
        logger.info(f"Definindo largura da coluna {column} para {width}")
        ws.set_column(column, column, width)

def protect(wb, ws, protect_config):
    if not isinstance(protect_config, dict):
        raise ValueError("'protect' deve ser um dicionário. Exemplo: {'password': 'senha'}")
    password = protect_config.get('password', None)
    options = protect_config.get('options', {})
    logger.info(f"Protegendo planilha com senha: {password}, opções: {options}")
    ws.protect(password, options)
