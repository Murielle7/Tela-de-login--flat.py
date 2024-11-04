import flet as ft

def main(page: ft.Page):
    page.title = "Tela de Login"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.bgcolor = "#E6E6FA"  # Fundo (lilás)

    # Adicionando o cabeçalho
    header = ft.Text("Atividade Avaliativa - Dev. Mobile", size=20, weight=ft.FontWeight.BOLD, color="#800080")

    # Adicionando a imagem centralizada
    image_path = "C:/Users/mferreirasantos/Pictures/images.png"
    image = ft.Image(src=image_path, width=200, height=200)  # Tamanho da imagem

    username = ft.TextField(label="Nome de Usuário", width=300)
    password = ft.TextField(label="Senha", password=True, can_reveal_password=True, width=300)

    # Estilização dos textos
    username.label_style = ft.TextStyle(color="#800080")  # Roxo usado nos textos
    password.label_style = ft.TextStyle(color="#800080")  

    def login_click(e):
        print(f"Login clicado, usuário: {username.value}, senha: {password.value}")

    login_button = ft.ElevatedButton(text="Login", on_click=login_click, style=ft.ButtonStyle(color="#800080"))  # Roxo

    # Adicionando o rodapé
    footer = ft.Text("Aluna: Murielle ----- 7° período", size=12, color="#800080")

    # Centralizando os elementos na página
    page.add(header, image, username, password, login_button, footer)

ft.app(target=main)
