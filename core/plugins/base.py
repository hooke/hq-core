class BasePlugin:
    def __init__(self, config):
        self.config = config

    def run(self):
        raise NotImplementedError("Plugin must implement the run() method.")
