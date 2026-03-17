from core.processor import CessaoProcessor


class RobotController:

    def __init__(self):
        self.processor = CessaoProcessor()

    def run(self):
        print("Controller iniciando execução")
        self.processor.executar()