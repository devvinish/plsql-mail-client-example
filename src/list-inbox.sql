select * from table(mail_client.get_mail_headers()) order by msg_number desc;
