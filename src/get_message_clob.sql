SELECT Mail_Client.Get_Message(1 /* specify message number */).get_bodypart_content_clob('0,1')
            FROM Dual;
            