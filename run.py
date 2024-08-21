import subprocess

def run_code(command):
    # Run the command and capture the output
    result = subprocess.run(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout, result.stderr

# Código 1
command1 = "C:\\Users\\jhennifer.nascimento\\nfs\\novo.py"  # Substitua pelo seu comando ou caminho do script
output1, error1 = run_code(command1)

# Código 2
command2 = "C:\\Users\\jhennifer.nascimento\\nfs\\novo.py"  # Substitua pelo seu comando ou caminho do script
output2, error2 = run_code(command2)

# Exibindo os resultados
print("Resultado do Código 1:")
print(output1)
if error1:
    print("Erros no Código 1:")
    print(error1)

print("\nResultado do Código 2:")
print(output2)
if error2:
    print("Erros no Código 2:")
    print(error2)
