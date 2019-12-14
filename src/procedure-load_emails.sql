CREATE OR REPLACE PROCEDURE LOAD_EMAILS IS
    
    CURSOR c_Inbox IS
      SELECT Msg_Number,
             Subject,
             Sender,
             Sender_Email,
             Sent_Date,
             Content_Type
        FROM TABLE(Mail_Client.Get_Mail_Headers())
       ORDER BY Msg_Number DESC;
  
    c_Clob    CLOB;
    b_blob    BLOB;
  
    t_Msg Mail_t;
  
    v_Partindex VARCHAR2(100);
  BEGIN
  
    Mail_Client.Connect_Server(p_Hostname => 'YOURMAILSERVER',
                               p_Port     => YOURPORT,
                               p_Protocol => Mail_Client.Protocol_Imap,
                               p_Userid   => 'USERID',
                               p_Passwd   => 'PASSWORD',
                               p_Ssl      => TRUE);
  
    Mail_Client.Open_Inbox;
  
    FOR c IN c_Inbox LOOP

     Dbms_Lob.Createtemporary(Lob_Loc => c_Clob,
                                   Cache   => TRUE,
                                   Dur     => Dbms_Lob.Call);

      Dbms_Lob.Createtemporary(Lob_Loc => b_blob,
                                   Cache   => TRUE,
                                   Dur     => Dbms_Lob.Call);
      
      IF Substr(c.Content_Type,
                   1,
                   9) = 'multipart' THEN
        v_Partindex := NULL;
        BEGIN
          SELECT Partindex
            INTO v_Partindex
            FROM TABLE(Mail_Client.Get_Message(c.Msg_Number).Get_Structure())
           WHERE Substr(Content_Type,
                        1,
                        9) = 'text/html';
        EXCEPTION
          WHEN OTHERS THEN
            NULL;
        END;
      
        IF v_Partindex IS NOT NULL THEN
        
          BEGIN
            SELECT Mail_Client.Get_Message(c.Msg_Number).Get_Bodypart_Content_Clob(v_Partindex)
              INTO c_Clob
              FROM Dual;
          EXCEPTION
            WHEN OTHERS THEN
              NULL;
          END;
          
                    BEGIN
            SELECT Mail_Client.Get_Message(c.Msg_Number).Get_Bodypart_Content_BLOB('1')
              INTO b_blob
              FROM Dual;
          EXCEPTION
            WHEN OTHERS THEN
              NULL;
          END;

        END IF;
        INSERT INTO mail_inbox
          (Msg_Number,
           Subject,
           Sent_Date,
           Sender_email,
           Body_Text,
           mail_attachment)
        VALUES
          (c.Msg_Number,
           c.Subject,
           c.Sent_Date,
           c.Sender_Email,
           c_Clob,
           b_blob);
      ELSIF Substr(c.Content_Type,
                   1,
                   9) = 'text/html' THEN
      
        BEGIN
          SELECT Mail_Client.Get_Message(c.Msg_Number).Get_Content_Clob()
            INTO c_Clob
            FROM Dual;
        EXCEPTION
          WHEN OTHERS THEN
            NULL;
        END;
        
                INSERT INTO mail_inbox
          (Msg_Number,
           Subject,
           Sent_Date,
           Sender_email,
           Body_Text)
        VALUES
          (c.Msg_Number,
           c.Subject,
           c.Sent_Date,
           c.Sender_Email
           c_Clob);

      END IF;
    END LOOP;
    COMMIT;
    Mail_Client.Close_Folder;
    Mail_Client.Disconnect_Server;
  
  EXCEPTION
    WHEN OTHERS THEN
      ROLLBACK;
    
      IF Mail_Client.Is_Connected() = 1 THEN
        Mail_Client.Close_Folder;
        Mail_Client.Disconnect_Server;
      END IF;
      RAISE;

  END LOAD_EMAILS;
  