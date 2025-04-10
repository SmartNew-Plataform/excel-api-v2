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
            set_rows(wb, ws, tags['set_row'])
        if 'set_column' in tags:
            logger.info(f"Aplicando set_column: {tags['set_column']}")
            set_columns(wb, ws, tags['set_column'])
        if 'protect' in tags:
            logger.info(f"Aplicando protect: {tags['protect']}")
            protect(wb, ws, tags['protect'])
    except Exception as e:
        logger.error(f"Erro ao aplicar tags: {str(e)}", exc_info=True)
        raise ValueError(f"Erro ao aplicar tags: {str(e)}")
    logger.info("add_tags concluído")