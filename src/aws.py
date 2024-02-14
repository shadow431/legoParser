import logging, boto3 # type: ignore

class aws:
    def __init__(self):
        self.logger = logging.getLogger('legoparser.aws')
    def get_ssm_parameter(self,parameter_name):
        ssm = boto3.client('ssm')
        parameter = ssm.get_parameter(Name=parameter_name, WithDecryption=True)
        return parameter['Parameter']['Value']