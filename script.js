function sleep(milliseconds) {
	var start = new Date().getTime();
	for (var i = 0; i < 1e7; i++) {
		if ((new Date().getTime() - start) > milliseconds) {
			break;
		}
	}
}


var Zname;
var Zemail;
var Zsenderemail;
var Zsubject;
var Zmessage;
var Zpassword;
var Data;

function getExcel() {
	var input = document.getElementById('fileinput');
	input.addEventListener('change', function () {
		readXlsxFile(input.files[0]).then(function (data) {
			Data=data;
		});
	});
}

function send()
{
	for (let i = 1; i < Data.length; i++) {
		console.log(Data[i][0] + " " + Data[i][1]);
		Zsenderemail = Data[i][1];
		alert(Zsenderemail);
		sendEmail();
	}
}

// function sendmail() {
// 	Zsenderemail = "arthurthomas7370@gmail.com";
// 	sendEmail();
// }

function sendEmail() {
	Zname = $('#Name').val();
	Zemail = $('#Sender').val();
	Zsubject = $('#Subject').val();
	Zmessage = $('#Message').val();
	Zpassword = $('#Password').val();

	Email.send({


		SecureToken: "fbf31702-bb7f-4a4e-9c1c-4ccf17ee777f",
		To: Zsenderemail,
		Host: "smtp.gmail.com",
		Username: Zemail,
		Password: Zpassword,
		From: Zemail,
		Subject: Zsubject,
		Body: "Name: " + Zname + "<br/>" + Zmessage


	}).then(
		message => {
			if (message == 'OK') {
				alert('Your mail has been send. Thank you for connecting.');
			}
			else {
				console.error(message);
				alert('There is error at sending message. ');
			}
		}
	);


}