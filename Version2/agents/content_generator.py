import ollama
import random

class ContentGeneratorAgent:
    def generate_content(self, prompt):
        try:
            response = ollama.generate(model="llama3.2", prompt=f"{prompt} Provide only the concise, complete text or numbered list (no introductory phrases, no formatting). Ensure 5 to 6 complete bullet points ending with full sentences, derived solely from the provided CSV data analysis.")
            return response['response'].strip()
        except Exception:
            return "Analysis failed due to error.\nCSV data could not be processed.\nPlease verify file integrity.\nContact support for assistance.\nThis is an error state."

    def split_into_bullets(self, text, min_points=5, max_points=6):
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if not lines or len(lines) < min_points:
            return [
                "Insufficient data in CSV.",
                "Analysis cannot be completed.",
                "Please check CSV content.",
                "No insights can be derived.",
                "This is an error message."
            ]
        num_points = random.randint(min_points, min(max_points, len(lines)))
        return lines[:num_points]