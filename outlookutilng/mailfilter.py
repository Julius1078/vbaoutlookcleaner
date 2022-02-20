class MailFilter:
    def __init__(self, name):
        self._name = name

    def filter_query(self):
        raise NotImplementedError

    def __str__(self) -> str:
        return '[Filter] "%s"' % self._name


class MultipleUsernameMailFilter(MailFilter):
    def __init__(self, maillist) -> None:
        self._maillist = maillist
        super().__init__('MultipleUsernameMailFilter')

    def filter_query(self):
        return "@SQL=" + " OR ".join([" ""http://schemas.microsoft.com/mapi/proptag/0x5D01001F"" = '{}'".format(mail) for mail in self._maillist.split(',')])
        # return ' OR '.join(['[SenderEmailAddress] = "{}"'.format(mail) for mail in self._maillist.split(',')])


class SubjectFilter(MailFilter):
    def __init__(self, subject, limit_time=True) -> None:
        self._subject = subject
        self._limit_time = limit_time
        super().__init__("SubjectFilter")

    def filter_query(self):
        if self._limit_time:
            return "@SQL=%last7days(""urn:schemas:httpmail:datereceived"")% AND ""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" like '{}'".\
                format(self._subject)
        return "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" like '{}'".\
            format(self._subject)


class BodyFilter(MailFilter):
    def __init__(self, body, limit_time=True) -> None:
        self._body = body
        self._limit_time = limit_time
        super().__init__("BodyFilter")

    def filter_query(self):
        if self._limit_time:
            return "@SQL=%last7days(""urn:schemas:httpmail:datereceived"")% AND ""urn:schemas:httpmail:textdescription"" ci_phrasematch '{}'".\
                format(self._body)
        return "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch '{}'".\
            format(self._body)
