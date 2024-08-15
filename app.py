import os
import smtplib
import pandas as pd

from flask import Flask, render_template, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_mail import Mail, Message



app = Flask(__name__)

base_directorio = os.path.abspath(os.path.dirname(__file__))
db_path = os.path.join(base_directorio, 'database/users.db')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'
db = SQLAlchemy(app)


app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = 'tolinsgabriel@gmail.com'
app.config['MAIL_PASSWORD'] = 'szcclwtuenddaqsi'
app.config['MAIL_DEFAULT_SENDER'] = 'tolinsgabriel.gmail.com'
mail = Mail(app)



class Users(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100))
    email = db.Column(db.String(150))
    telefone = db.Column(db.String(20))
    endereco = db.Column(db.String(200))
    conhecimento = db.Column(db.String(150))

with app.app_context():
    db.create_all()
    db.session.commit()


def email_conf(destinatario, nome):
    msg = Message(
        subject='Confirmação de Cadastro na Santi Clinic',
        recipients=[destinatario],
        body=f'Seja Bem-vindo {nome},\n\nObrigado por se cadastrar na Santiclinic. Seu cadastro foi efetuado com sucesso!\n\nAtenciosamente,\nEquipe Santi Clinic.'
    )
    try:
        msg = Message(
            subject='Confirmação de Cadastro na Santi Clinic',
            recipients=[destinatario],
            body=f'Seja Bem-vindo {nome},\n\nObrigado por se cadastrar na Santiclinic. Seu cadastro foi efetuado com sucesso!\n\nAtenciosamente,\nEquipe Santi Clinic'
        )
        mail.send(msg)
        print('E-mail enviado com sucesso')
    except smtplib.SMTPConnectError as e:
        print(f'Erro de conexão SMTP: {e}')
    except smtplib.SMTPAuthenticationError as e:
        print(f'Erro de autenticação SMTP: {e}')
    except Exception as e:
        print(f'Erro ao enviar e-mail: {e}')


@app.route('/')
def home():
    return render_template('home.html')

@app.route('/add_user', methods=['POST'])
def criar():
    nome = request.form['nome']
    email = request.form['email']
    telefone = request.form['telefone']
    endereco = request.form['endereco']
    conhecimento = request.form['conhecimento']


    user = Users(nome=nome, email=email, telefone=telefone, endereco=endereco, conhecimento=conhecimento)
    db.session.add(user)
    db.session.commit()
    email_conf(email, nome)
    return redirect(url_for('home'))

@app.route('/exportar')
def exportar_dados():
    users = Users.query.all()

    data = {
        'ID': [user.id for user in users],
        'Nome': [user.nome for user in users],
        'Email': [user.email for user in users],
        'Telefone':[user.telefone for user in users],
        'Endereço':[user.endereco for user in users],
        'Conhecimento':[user.conhecimento for user in users],
    }

    df = pd.DataFrame(data)

    caminho = os.path.join(base_directorio, 'users.xlsx')

    df.to_excel(caminho, index=False)

    return  send_file(caminho, as_attachment=True, download_name='users.xlsx')





if __name__ == '__main__':
    app.run(debug=True)

