import argparse

def main():
    
    '''
    #subparsers = parser.add_subparsers(dest='mode', required=True)  # Python 3.7+
    # CLI Mode
    cli_parser = subparsers.add_parser('cli', help='Command-line mode')
    cli_parser.add_argument('--path', required=True, help='File path for CLI')
    # GLI Mode
    gli_parser = subparsers.add_parser('gli', help='Graphical mode')
    '''

    if args.mode == 'cli':
        print(f"CLI Mode - Path: {args.path}")  # Replace with your actual function
        try:
            from parts.cli import mymain # Import inside the function to avoid circular imports
            mymain(args.path)
            #print('success')
        except:
            print(f"Error importing cli.py:")
    elif args.mode == 'gui':
        try:
            import parts.gui # Import inside the function to avoid circular imports
            #ymain(args.path)
            print('success')
        except:
            print(f"Error in gui.py:")

if __name__ == "__main__":
    
    parser = argparse.ArgumentParser(description="Run in CLI or GUI mode.")
    parser.add_argument('--mode', choices=['cli', 'gui'], required=True, help='Specify the mode' )
    parser.add_argument('--path', help='enter the path')    
    args = parser.parse_args()
    print(args.mode)
    main()