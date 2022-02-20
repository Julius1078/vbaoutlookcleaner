import win32com.client
from .mailrule import Rule
from .mailfilter import MultipleUsernameMailFilter
from typing import Sequence

class MailCleaner:

    def __init__(self, filter_personal_name = None) -> None:        
        self._filterNonPersonal = None
        if (filter_personal_name):
            self._filterNonPersonal = Rule('Non Personal E-mail', MultipleUsernameMailFilter(filter_personal_name))
        self._rules =[]
            #Initialize outlook
        self._inbox_folder = win32com.client.Dispatch(
                "Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)            

    def addrule(self, rule : Rule):
        self._rules.append(rule)

    def addrules(self, rulelist: Sequence[Rule]):
        for rule in rulelist:
            self.addrule(rule)

    def cleanAll(self):
        for rule in self._rules:
            rule.applyRule(self._inbox_folder)
        if self._filterNonPersonal:
            self._filterNonPersonal.applyRule(self._inbox_folder)

    def cleanMessage(self,message):
        for rule in self._rules:
            rule.applyRuletoMessage(message)


