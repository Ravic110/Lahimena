"""Cache partage des distances KM_MADA.

Isole dans son propre module pour que les tables de reference puissent
l'invalider sans dependre de utils.excel_handler, qui les importe.

Le dictionnaire est mute sur place plutot que rebindé : tout module l'ayant
importe voit donc la meme instance, y compris apres invalidation.
"""

_KM_MADA_CACHE_TTL_SECONDS = 10.0

_KM_MADA_CACHE = {
    "path": None,
    "mtime": None,
    "loaded_at": 0.0,
    "rows": [],
    "lookup": {},
}


def _invalidate_km_mada_cache():
    """Vide le cache des distances."""
    _KM_MADA_CACHE["path"] = None
    _KM_MADA_CACHE["mtime"] = None
    _KM_MADA_CACHE["loaded_at"] = 0.0
    _KM_MADA_CACHE["rows"] = []
    _KM_MADA_CACHE["lookup"] = {}
