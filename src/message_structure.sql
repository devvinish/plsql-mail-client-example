/* specify message number for get_message() to see its structure */
select * from table(mail_client.get_message(1).get_structure());
