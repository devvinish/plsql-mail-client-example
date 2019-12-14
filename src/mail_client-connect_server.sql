begin
  mail_client.connect_server(
    p_hostname => 'YourMailServer.com',
    p_port     => YourPortIntegerValue,
    p_protocol => mail_client.protocol_IMAP, -- or mail_client.protocol_POP3
    p_userid   => 'YourUserID',
    p_passwd   => 'YourPassword',
    p_ssl      => true -- true or false depends on your mailbox
  );

  mail_client.open_inbox;
  dbms_output.put_line('Mailbox successfully opened.');
  dbms_output.put_line('The INBOX folder contains '||mail_client.get_message_count||' messages.');
end;
/