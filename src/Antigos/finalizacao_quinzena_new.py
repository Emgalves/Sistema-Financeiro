import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from datetime import datetime
from dateutil.relativedelta import relativedelta
import calendar
import os
import openpyxl
import sys
from pathlib import Path

# Ajustar o path do sistema
current_dir = Path(__file__).resolve().parent
project_root = current_dir.parent
if str(project_root) not in sys.path:
    sys.path.append(str(project_root))

# Garantir que o diretório src está no path
src_dir = project_root / 'src'
if str(src_dir) not in sys.path:
    sys.path.append(str(src_dir))

# Importar GestaoTaxasFixas
try:
    # Tentar importação absoluta primeiro
    from Sistema_Entrada_Dados import GestaoTaxasFixas
except ImportError:
    try:
        # Tentar com prefixo src
        from src.Sistema_Entrada_Dados import GestaoTaxasFixas
    except ImportError:
        # Última tentativa: importar do diretório atual
        from .Sistema_Entrada_Dados import GestaoTaxasFixas

# Importar utilitários
try:
    # Tentar importação absoluta primeiro
    from config.utils import (
        validar_data,
        validar_data_quinzena,
        formatar_moeda,
        ARQUIVO_CLIENTES,
        ARQUIVO_CONTROLE,
        PASTA_CLIENTES,
        BASE_PATH
    )
except ImportError:
    try:
        # Tentar com prefixo src
        from src.config.utils import (
            validar_data,
            validar_data_quinzena,
            formatar_moeda,
            ARQUIVO_CLIENTES,
            ARQUIVO_CONTROLE,
            PASTA_CLIENTES,
            BASE_PATH
        )
    except ImportError:
        # Última tentativa: importar do diretório atual
        from .config.utils import (
            validar_data,
            validar_data_quinzena,
            formatar_moeda,
            ARQUIVO_CLIENTES,
            ARQUIVO_CONTROLE,
            PASTA_CLIENTES,
            BASE_PATH
        )