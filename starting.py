from app import kick_start as app_kick
from app_request import kick_start as app_req_kick


def app_kick_fn():
    if __name__ == '__main__':
        for mats in range(20620, 206275):
            app_kick(str(mats), str(mats))
        print('Done')
        while True:
            pass


def app_req_kick_fn():
    if __name__ == '__main__':
        while True:
            mats = input("Which profile to fetch ? >> ")
            app_req_kick(str(mats), str(mats))
            print('---------------Done-------------------')


operation = int(input("Request (1) or Multiple? (2) "))
if operation == 2:
    app_kick_fn()
elif operation == 1:
    app_req_kick_fn()
else:
    print("Invalid input")
