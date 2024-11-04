import json
import os

def cadastrar_aluno(alunos):
    nome = input("Digite o nome do aluno: ")
    idade = input("Digite a idade do aluno: ")
    telefone = input("Digite o telefone do aluno: ")
    
    aluno = {
        "nome": nome,
        "idade": idade,
        "telefone": telefone
    }
    
    alunos.append(aluno)
    salvar_alunos(alunos)
    print("Aluno cadastrado com sucesso!")

def salvar_alunos(alunos):
    with open('aluno.db.json', 'w', encoding='utf-8') as f:
        json.dump(alunos, f, ensure_ascii=False, indent=4)

def mostrar_alunos(alunos):
    if alunos:
        print("\nAlunos cadastrados:")
        for idx, aluno in enumerate(alunos, start=1):
            print(f"{idx}. Nome: {aluno['nome']}, Idade: {aluno['idade']}, Telefone: {aluno['telefone']}")
    else:
        print("Nenhum aluno cadastrado.")

def buscar_aluno(alunos):
    termo = input("Digite o nome (ou parte do nome) que deseja buscar: ")
    resultados = [aluno for aluno in alunos if aluno['nome'].startswith(termo)]
    
    if resultados:
        print("\nResultados da busca:")
        for aluno in resultados:
            print(f"Nome: {aluno['nome']}, Idade: {aluno['idade']}, Telefone: {aluno['telefone']}")
    else:
        print("Nenhum aluno encontrado com esse termo.")

def carregar_alunos():
    if os.path.exists('aluno.db.json'):
        with open('aluno.db.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def main():
    alunos = carregar_alunos()
    
    while True:
        print("\nMenu:")
        print("1. Cadastrar Aluno")
        print("2. Mostrar Alunos")
        print("3. Buscar Aluno")
        print("4. Sair")
        
        opcao = input("Escolha uma opção: ")
        
        if opcao == '1':
            cadastrar_aluno(alunos)
        elif opcao == '2':
            mostrar_alunos(alunos)
        elif opcao == '3':
            buscar_aluno(alunos)
        elif opcao == '4':
            print("Saindo do programa.")
            break
        else:
            print("Opção inválida. Tente novamente.")

if __name__ == "__main__":
    main()
