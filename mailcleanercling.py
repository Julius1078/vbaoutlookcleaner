from outlookutilng import MailCleaner,  OutlookRulesConfigParser

parser = OutlookRulesConfigParser('./config.ini')

mc = MailCleaner()
mc.addrules(parser.get_rules())
mc.cleanAll()

