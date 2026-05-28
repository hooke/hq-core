import importlib

class PipelineEngine:
    def __init__(self, config):
        self.config = config

    def run_plugin(self, plugin_module, plugin_class):
        module = importlib.import_module(f"core.plugins.{plugin_module}")
        plugin = getattr(module, plugin_class)(self.config)
        plugin.run()
