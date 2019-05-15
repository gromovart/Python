import socket
import threading
import sys

sock = socket.socket()
sock.bind(('localhost', 3003))
sock.listen(10)
conn = []

def request_to_read():
    while True:
        for i in range(len(conn)):
            try:
                request = conn[i].recv(1024)
                if request:
                    print(request.decode())
            except socket.error as e:
                if e.errno == 10053:
                    conn.pop(i)
                    print("Connected users:", len(conn))
                else:
                    raise
def send_msg():
    while True:
        global conn
        message = input()
        if message:
            for i in range(len(conn)):
                conn[i].send(message.encode())

def conn_users():
    while True:
        global conn
        conn.append(sock.accept()[0])
        print("Connected users:", len(conn))


# init threads
t1 = threading.Thread(target=request_to_read)
t2 = threading.Thread(target=send_msg)
t3 = threading.Thread(target=conn_users)

# start threads
t1.start()
t2.start()
t3.start()