import configparser

def get_config_variables(filepath=str):
    """ Reads a configuration file and returns variables in a dictionary key, value pairs. The
    filepath must be given as a raw string, e.g. r'C:\path\to\file.ini' """
    config = configparser.ConfigParser()
    config.read(filepath)

    variables = {}

    for section in config.sections():
        for key, value in config[section].items():
            variables[key] = value
        #print(variables)
    return variables
#get_variables("D:\\Background_Verification_Bot\\config.ini")