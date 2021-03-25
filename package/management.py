import lib
import fatures as f

def main():
    a = lib.Login()
    shenfen = a.login_interface()
    b = lib.Denglu()
    if shenfen=='admin':
        b.admin_login()
    elif shenfen=='tourist':
        b.tourist_login()

if __name__ == '__main__':
    main()
