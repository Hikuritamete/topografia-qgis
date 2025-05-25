def classFactory(iface):
    from .plugin import TopografiaPlugin
    return TopografiaPlugin(iface)