#!/usr/bin/python

PROJECT_NAME = 'Top 1000 Words in Gilaki'
DB_DIR = 'words.db'
RELEASE_DIR = "./" + PROJECT_NAME.replace(" ", "-")
START_RANGE = 1
END_RANGE = 1001
VERSION = "2.0.0"
COMMIT_MESSAGE = "fix typo " + VERSION

if __name__ == '__main__':
    print("You can not run this file directly.")
