def get_folder(root_folder, subfolder_name, createFolder=True):
    if (not createFolder):
        return root_folder.Folders[subfolder_name]
    else:
        try:
            return root_folder.Folders[subfolder_name]
        except:
            return root_folder.Folders.Add(subfolder_name)


class Rule:
    def __init__(self, rulename, root_folder, filter, make_subfolder=False) -> None:
        self._rulename = rulename
        self._root_folder = root_folder
        self._filter = filter
        self._make_subfolder = make_subfolder

    def applyRule(self, mailfolder) -> None:
        dst_folder = mailfolder
        for folder in self._root_folder.split('/'):
            dst_folder = get_folder(dst_folder, folder)
        #dst_folder = get_folder(mailfolder, self._root_folder)
        # Filter items on mailfolder folder
        filtered_elements = mailfolder.Items.Restrict(self._filter.filter_query())
        while (len(filtered_elements) > 0):
            for element in filtered_elements:
                if element.MessageClass not in ('IPM.Outlook.Recall', 'IPM.Schedule.Meeting'):
                    try:
                        # Create subfolder if required
                        if self._make_subfolder:
                            # Check Email 
                            if element.SenderEmailType == 'EX':
                                sender = element.Sender.GetExchangeUser().PrimarySmtpAddress
                            else:
                                sender = element.SenderEmailAddress
                            sub_folder = get_folder(dst_folder, sender)
                            print('[{}] Moving element "{}" to {}'.format(
                                self._rulename, element.Subject, sub_folder))
                            element.move(sub_folder)
                        else:                            
                            print('[{}] Moving element "{}" to {}'.format(
                                self._rulename, element.Subject, dst_folder))
                            element.move(dst_folder)                            
                    except Exception as e:
                        print('Error ocurred {}'.format(e))
            filtered_elements = mailfolder.Items.Restrict(self._filter.filter_query())      