"""
Compatibilidad: el robot ahora está dividido en un archivo por paso y se ejecuta con
lanzador.py. Este archivo solo llama al lanzador para que `python rpa_dian.py` siga funcionando.
"""

from lanzador import main

if __name__ == "__main__":
    main()
