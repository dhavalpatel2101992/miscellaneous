import paramiko
import sys
from cryptography.fernet import Fernet
import os
import logging
wrk_dir=r"E:\Control-M\tabprd\FPM\AP_Forecast_Model\\"
os.chdir(wrk_dir)
key = b'icX0uAcHik9UCEgzTY3jP2_KhbEfXGAZucdSX3sbQMQ='
cipher_suite = Fernet(key)
def get_logger(script,logname,loglevel="DEBUG"):
    # Default the logging level to INFO if no LOGGING_LEVEL param defined
    level = os.environ.get("AP_FORECASTING_LOGGING_LEVEL")
    if level == None:
        level = loglevel
    # Create/Get a custom logger
    logger = logging.getLogger(script,)
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(logname)
    # Set the level
    if level:
        fh.setLevel(logging.getLevelName(level))
    else:
        fh.setLevel(logging.DEBUG)

    # create console handler with a higher log level
    ch = logging.StreamHandler()
    # ch.setLevel(logging.ERROR)
    ch.setLevel(logging.DEBUG)

    # Create formatter and add it to handler
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger
#default logger
logger = get_logger('AP_Forecast_Model','AP_Forecast_Model.log')
print = logger.debug

class SshClient:
    "A wrapper of paramiko.SSHClient"
    TIMEOUT = 4

    def __init__(self, host, port, username, password):
        self.username = username
        self.password = password
        self.client = paramiko.SSHClient()
        self.client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        self.client.connect(host, port, username=username, password=password, pkey=None, timeout=self.TIMEOUT)

    def close(self):
        if self.client is not None:
            self.client.close()
            self.client = None

    def execute(self, command, sudo=False):
        feed_password = False
        if sudo and self.username != "root":
            # command = "sudo -S -p '' %s" % command
            feed_password = self.password is not None and len(self.password) > 0
        stdin, stdout, stderr = self.client.exec_command(command)
        if feed_password:
            stdin.write(self.password + "\n")
            stdin.flush()
        return {'out': stdout.readlines(),
                'err': stderr.readlines(),
                'retval': stdout.channel.recv_exit_status()}

try:
    file = open("password.txt","r")
    password = cipher_suite.decrypt(bytes(file.read(), 'utf-8')).decode('utf-8')
    print('retrieved password from password.txt')
except:
    print("password.txt does not exist")
    exit(0)

if __name__ == "__main__":
    try:
        hostname='ebimlprd1'
        username='fpm_ds'
        client = SshClient(host=hostname, port=22, username=username, password=password)
        print("user {0} connected to : {1}".format(username,hostname))
    except:
        print("Unable to connect to the server : Server is down or Password is incorrect")
        exit(0)
    try:
        # cmd="bash /local/mnt/workspace/fpm_ds/corpfin-ap-forecast/account_payables.sh"
        print("{0} is getting triggered ...".format(cmd))
        ret = client.execute(cmd, sudo=True)
        print('{0} {1} {2}'.format("  ".join(ret["out"]), "  E ".join(ret["err"]), ret["retval"]))
        print('execution successful ...')
    except Exception as e:
        print('unknown error occurred - {0}'.format(str(e)))
    finally:
      client.close()