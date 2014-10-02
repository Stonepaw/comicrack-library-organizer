import System

import settings


def main():

    profiles = settings.get_profiles()

    print profiles
    print type(profiles)
    settings.save_profiles(profiles)
    System.Console.ReadLine()

if __name__ == "__main__":
    main()