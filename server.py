from flask import Flask
import subprocess

app = Flask(__name__)

@app.route('/run-ilogic', methods=['POST'])
def run_ilogic():
    # Replace with the full path to your .exe file
    subprocess.Popen([r"D:\new_proj\Inventor_final\Inventor_final\bin\x64\Debug\Inventor_final.exe"])
    return "iLogic Rule Triggered", 200

if __name__ == '__main__':
    app.run(debug=True)
