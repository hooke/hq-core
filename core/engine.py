import importlib

class PipelineEngine:
    def __init__(self, config):
        self.config = config

    def run_plugin(self, module_path, plugin_class):
        """
        module_path: i.e. 'core.plugins.p01_ingester' 
                     or 'custom_plugins.p02_secret_generator'
        """
        # Dynamically load module by the path
        module = importlib.import_module(module_path)
        # Get plugin class & create instance
        plugin = getattr(module, plugin_class)(self.config)
        # run 
        plugin.run()
