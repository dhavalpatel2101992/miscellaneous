from cryptography.fernet import Fernet
key = b'icX0uAcHik9UCEgzTY3jP2_KhbEfXGAZucdSX3sbQMQ='
cipher_suite = Fernet(key)
password = input('Enter Password:')
secure_password = cipher_suite.encrypt(bytes(password, 'utf-8')).decode('utf-8')
file = open(r"password.txt","w")
file.write(secure_password)
file.close()
print('Saved Encrypted Password in password.txt')

