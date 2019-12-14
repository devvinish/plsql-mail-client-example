Create or Replace PROCEDURE Delete_Mail_Msg(i_Msg_Number IN NUMBER) IS
     
t_Msg Mail_t;

BEGIN

Mail_Client.Connect_Server(p_Hostname => 'YourMailServer',
                           p_Port     => MailServerPort,
                           p_Protocol => Mail_Client.Protocol_Imap,
                           p_Userid   => 'username',
                           p_Passwd   => 'password',
                           p_Ssl      => TRUE);

Mail_Client.Open_Inbox;

t_Msg := Mail_Client.Get_Message(i_Msg_Number);
t_Msg.Mark_Deleted();

Mail_Client.Expunge_Folder;
Mail_Client.Close_Folder;
Mail_Client.Disconnect_Server;

EXCEPTION
     WHEN OTHERS THEN
       IF Mail_Client.Is_Connected() = 1 THEN
         Mail_Client.Close_Folder;
         Mail_Client.Disconnect_Server;
       END IF;
       Raise;
   END Delete_Mail_Msg;
   