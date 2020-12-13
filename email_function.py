import smtplib


def email_sent_func(to_, subject_, message_, email_, password_):
	# print(to_, subject_, message_, email_, password_)
	s = smtplib.SMTP("smtp.gmail.com", 587)
	s.starttls()  # THE TRANSPORT LAYER
	s.login(email_, password_)
	msg = "Subject: {}\n\n{}".format(subject_, message_)
	s.sendmail(email_, to_, message_)
	x = s.ehlo()

	if x[0]==250:
		return "s"
	else:
		return "f"
	s.close()