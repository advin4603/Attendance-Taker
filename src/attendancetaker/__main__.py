from attendancetaker.app import main
import os

if __name__ == '__main__':
    if not os.path.isdir("Tracebacks"):
        os.mkdir("Tracebacks")
    main()
