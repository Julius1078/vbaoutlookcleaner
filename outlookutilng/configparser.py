from configparser import RawConfigParser
from .mailfilter import *
from .mailrule import Rule


class OutlookRulesConfigParser():

    def __init__(self, configfile) -> None:
        self.config = RawConfigParser()
        self.config.read(configfile,encoding='utf-8')
        # General Data (Fulle Name and )
        self.personalname = self.config['General']['personal_name']
        # Process Rules. One Section for each Folder with its own config
        self._rules = []
        for section in [section for section in self.config.sections() if section.endswith('_Folder')]:
            config = {item[0]: item[1] for item in self.config[section].items(
            ) if item[0] not in ['destinationfolder', 'ruletype']}
            try:
                self._rules.append(Rule(section[:-7], self.config[section]['DestinationFolder'], globals()[self.config[section]['RuleType']](**config),
                                        True if self.config[section]['RuleType'] in ['MultipleUsernameMailFilter'] else False))
            except Exception as e:
                print("Error creation rule {} of type {}. Error: {}",
                      section,
                      self.config[section]['RuleType'],
                      e
                      )

    def get_rules(self):
        return self._rules
