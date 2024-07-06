# Script Populis

### Uso

- Baixar o repositório como .zip
- Editar o json com as informações

        Ex:

        {
        "cpf": "12345678",
        "password": "12313123",
        "horarios": {
            "Entrada1": "13:00",
            "Saida1": "15:40",
            "Entrada2": "15:55",
            "Saida2": "19:00"
        }
        }
- Abir a pasta com o terminal

        pip install -r requirements.txt

    #

        python Script.py


# Como funciona:

Uma planilha é gerada dentro da pasta para armzenar os horários salvos no Json (somente para fins de controle)

O código monitora os horários salvos e executa a ação no horário definido

Após executar as 4 ações (Entrada1, Saida1, Entrada2, Saida2) o código encerra.
