import subprocess

try:
    output = subprocess.check_output("p4 info", shell=True, stderr=subprocess.STDOUT)
    text = output.decode(encoding='UTF-8')
    print(text)
except subprocess.CalledProcessError as e:
    print("Error Code:", e.returncode)
    print("Output:", e.output.decode(encoding='UTF-8'))