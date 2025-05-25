def classFactory(iface):
    from .topografia import TopografiaPlugin
    return TopografiaPlugin(iface)