from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import declarative_base, sessionmaker

# Definindo a base
Base = declarative_base()

# Definindo o modelo
class User(Base):
    __tablename__ = 'users'
    
    id = Column(Integer, primary_key=True)
    name = Column(String)
    age = Column(Integer)

    def __repr__(self):
        return f"<User(name='{self.name}', age={self.age})>"

# Criando o banco de dados
engine = create_engine('sqlite:///example.db')  # Usa um banco de dados SQLite
Base.metadata.create_all(engine)

# Criando uma sessão
Session = sessionmaker(bind=engine)
session = Session()

# Adicionando múltiplos usuários
while True:
    name = input("Informe o nome do usuário (ou 'sair' para finalizar, 'listar' para mostrar todos os usuários): ")
    if name.lower() == 'sair':
        break
    if name.lower() == 'listar':
        # Consultando e exibindo usuários cadastrados
        users = session.query(User).all()
        print("Usuários cadastrados:")
        for user in users:
            print(user)
        continue  # Volta para o início do loop

    age = input("Informe a idade do usuário: ")
    
    try:
        age = int(age)  # Converte a idade para um inteiro
        new_user = User(name=name, age=age)
        session.add(new_user)
        session.commit()
    except ValueError:
        print("Por favor, insira uma idade válida.")

# Fechando a sessão
session.close()
