SELECT Mail_Client.Get_Message(1 /* specify message number */).get_bodypart_content_varchar2('0,0')
             FROM Dual;
             