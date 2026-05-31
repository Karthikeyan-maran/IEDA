# llm_worker.py

from PyQt5.QtCore import QThread, pyqtSignal
import ollama

class LLMWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, prompt, dataframes_info):
        super().__init__()
        self.prompt = prompt
        self.dataframes_info = dataframes_info

    def run(self):
        try:
            context = f"Here's information about the loaded Excel dataframes:\n{self.dataframes_info}\n"
            full_prompt = f"{context}\nUser Request: {self.prompt}"
            response = ollama.chat(model='phi3', messages=[{'role': 'user', 'content': full_prompt}])
            generated_text = response['message']['content']
            self.finished.emit(generated_text)
        except Exception as e:
            self.error.emit(
                f"Error during LLM call: {str(e)}\n"
                "Please ensure Ollama is running and the 'phi3' model is downloaded."
            )
