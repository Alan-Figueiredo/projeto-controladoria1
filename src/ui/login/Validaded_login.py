import json

class Validaded_login:
    def __init__(self,user: str, passwd: str):
        self.user = user
        self.passwd = passwd
        self.db = None

    def auth(self):
        if self.passwd == self.readDB():
            return True
        else:
            False


    def readDB(self,user_name:str, user_passwd: str):
        with open("db.json","r", encoding="utf-8") as file:
            db_json = json.load(file)
            temp_senha = []
           
            for s in db_json:
                if s["passwd"] == user_passwd and s["user"] == user_name:
                    temp_senha = s["passwd"]

        return temp_senha

