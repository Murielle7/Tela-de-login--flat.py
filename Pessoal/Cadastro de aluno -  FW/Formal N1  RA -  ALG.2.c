#include <stdio.h>
#include <stdlib.h>
#include <string.h>

typedef struct Cadastro {
    char nome[100];
    char cpf[12];
    char telefone[15];
    char email[100];
} Cadastro;

typedef struct No {
    Cadastro dados;
    struct No *esquerda;
    struct No *direita;
} No;

// Função para criar um novo nó
No* criar_no(Cadastro cadastro) {
    No* novo_no = (No*)malloc(sizeof(No));
    novo_no->dados = cadastro;
    novo_no->esquerda = NULL;
    novo_no->direita = NULL;
    return novo_no;
}

// Função para inserir um nó na árvore
No* inserir_no(No* raiz, Cadastro cadastro) {
    if (raiz == NULL) {
        return criar_no(cadastro);
    }

    if (strcmp(cadastro.cpf, raiz->dados.cpf) < 0) {
        raiz->esquerda = inserir_no(raiz->esquerda, cadastro);
    } else {
        raiz->direita = inserir_no(raiz->direita, cadastro);
    }

    return raiz;
}

// Função para buscar um nó na árvore
No* buscar_no(No* raiz, char cpf[]) {
    if (raiz == NULL || strcmp(cpf, raiz->dados.cpf) == 0) {
        return raiz;
    }

    if (strcmp(cpf, raiz->dados.cpf) < 0) {
        return buscar_no(raiz->esquerda, cpf);
    } else {
        return buscar_no(raiz->direita, cpf);
    }
}

// Função para encontrar o nó com o menor valor (usado na remoção)
No* encontrar_minimo(No* no) {
    while (no && no->esquerda != NULL) {
        no = no->esquerda;
    }
    return no;
}

// Função para remover um nó da árvore
No* remover_no(No* raiz, char cpf[]) {
    if (raiz == NULL) {
        return raiz;
    }

    if (strcmp(cpf, raiz->dados.cpf) < 0) {
        raiz->esquerda = remover_no(raiz->esquerda, cpf);
    } else if (strcmp(cpf, raiz->dados.cpf) > 0) {
        raiz->direita = remover_no(raiz->direita, cpf);
    } else {
        // Nó encontrado, agora precisa ser removido
        if (raiz->esquerda == NULL) {
            No* temp = raiz->direita;
            free(raiz);
            return temp;
        } else if (raiz->direita == NULL) {
            No* temp = raiz->esquerda;
            free(raiz);
            return temp;
        }

        No* temp = encontrar_minimo(raiz->direita);
        raiz->dados = temp->dados;
        raiz->direita = remover_no(raiz->direita, temp->dados.cpf);
    }

    return raiz;
}

// Função para imprimir os cadastros em ordem crescente de CPF
void imprimir_arvore(No* raiz) {
    if (raiz != NULL) {
        imprimir_arvore(raiz->esquerda);
        printf("Nome: %s, CPF: %s, Telefone: %s, Email: %s\n", raiz->dados.nome, raiz->dados.cpf, raiz->dados.telefone, raiz->dados.email);
        imprimir_arvore(raiz->direita);
    }
}

int main() {
    No* raiz = NULL;
    int opcao;
    char cpf[12];
    Cadastro cadastro;

    do {
        printf("\n1. Inserir\n2. Buscar\n3. Remover\n4. Imprimir\n5. Sair\nEscolha uma opcao: ");
        scanf("%d", &opcao);

        switch (opcao) {
            case 1:
                printf("Nome: ");
                scanf(" %[^\n]", cadastro.nome);
                printf("CPF: ");
                scanf("%s", cadastro.cpf);
                printf("Telefone: ");
                scanf("%s", cadastro.telefone);
                printf("Email: ");
                scanf("%s", cadastro.email);
                raiz = inserir_no(raiz, cadastro);
                break;
            
            case 2:
                printf("Digite o CPF para buscar: ");
                scanf("%s", cpf);
                No* encontrado = buscar_no(raiz, cpf);
                if (encontrado != NULL) {
                    printf("Cadastro encontrado: Nome: %s, CPF: %s, Telefone: %s, Email: %s\n", encontrado->dados.nome, encontrado->dados.cpf, encontrado->dados.telefone, encontrado->dados.email);
                } else {
                    printf("Cadastro nao encontrado.\n");
                }
                break;
            
            case 3:
                printf("Digite o CPF para remover: ");
                scanf("%s", cpf);
                raiz = remover_no(raiz, cpf);
                break;
            
            case 4:
                imprimir_arvore(raiz);
                break;
            
            case 5:
                printf("Saindo...\n");
                break;

            default:
                printf("Opcao invalida!\n");
        }
    } while (opcao != 5);

    return 0;
}
