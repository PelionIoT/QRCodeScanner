// By A. Sean Kessel

/* Make sure prompt, xlsx, chalk, colors, and cli-colors are installed. 
If not, here are the commands to install everything needed to run the program:
npm install prompt
npm install xlsx
npm install chalk
npm install colors
npm install cli-color
*/
//it begins by prompting the user. The user scans the QR Code on the product
console.log("Please scan the QR code on the item being shipped");
var prompt = require('prompt');
prompt.get('number', function(err, result){
if (err) {
	console.log(err)
} else {

/* Below an excel file found in /Users/Kessel/Documents/Example1.xlsx is accessed.
The Excel file contains the addresses of the customer who bought the product originally scanned
*/

//CHANGE THE ADDRESS OF THE WORKBOOK FOR EACH COMPUTER

if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('/home/sean/Desktop/ShippingAndProductInformation.xlsx');

var first_sheet_name = workbook.SheetNames[0];
var wrksht = workbook.Sheets[first_sheet_name];

/* Below the program searches through the "B" column of the Excel file, searching for a value
that equals the product's QR code value
*/

// The counter ensures the right part of the code gets executed at the right stages. 
var counter = 0;


var i=3;
	while(counter===0) {
		var addressOfCell = 'B'+i ;
		var cell = wrksht[addressOfCell];
		
			var value_of_cell = cell.v;
		
			if (result.number == value_of_cell){
			/* Once there is a match, the program gets the values of the address,
			first name, last name, and random number associated with the product code
			which are found in Columns D, F, G, and I respectively.
			*/
			
			var cell_of_streetAddressLine1 = wrksht['D'+i];
				var streetAddressLine1 = cell_of_streetAddressLine1.v;
			var add = i+1;
			var cell_of_streetAddressLine2 = wrksht['D'+ add];
				var streetAddressLine2 = cell_of_streetAddressLine2.v;
			var cell_of_city = wrksht['F'+i];
				var city = cell_of_city.v;
			var cell_of_state = wrksht['G'+i];
				var state = cell_of_state.v;
			var cell_of_zipcode = wrksht['H'+i];
				var zipcode = cell_of_zipcode.v;
			var cell_of_country = wrksht['I'+i];
				var country = cell_of_country.v;
			var cell_firstName = wrksht['J'+i];
				var firstName = cell_firstName.v;
			var cell_lastName = wrksht['K'+i];
				var lastName = cell_lastName.v;
			var cell_randomNumber = wrksht['M'+i];
				var randomNumber = cell_randomNumber.v;

			/* Then, it writes a .zpl file. It encodes the scanned number on the product into
			a QR code, and adds the address of the clients.
			*/

			//CHANGE THE ADDRESS OF THE LABEL GETS SENT TO FOR EACH COMPUTER

			var fs = require('fs');
			var stream = fs.createWriteStream("/home/sean/Desktop/QRLabel.zpl");
			stream.once('open', function(fd) {
	  		stream.write("^XA \n");
	  		stream.write("\n");
			stream.write("^FX Top section with company logo, name and address. \n");
			stream.write("^CF0,60 \n");
			stream.write("^FO50,50^GB100,100,100^FS \n");
			stream.write("^FO75,75^FR^GB100,100,100^FS \n");
			stream.write("^FO88,88^GB50,50,50^FS \n");
			stream.write("^FO220,50^FD WigWag, Inc.^FS \n");
			stream.write("^CF0,40 \n");
			stream.write("^FO220,100^FD1000 4009 Banister Ln #200^FS \n");
			stream.write("^FO220,135^FD Austin TX 38102^FS \n");
			stream.write("^FO220,170^FDUnited States (USA)^FS \n");
			stream.write("^FO50,250^GB700,1,3^FS \n");
			stream.write("\n");
			stream.write("^FX Second section with recipient address and permit information. \n");
			stream.write("^CFA,30 \n");
			stream.write("^FO50,300^FD"+ firstName + " " + lastName +"^FS \n");
			stream.write("^FO50,340^FD302" + streetAddressLine1 +"^FS \n");
			stream.write("^FO50,380^FD"+ streetAddressLine2 +"^FS \n");
			stream.write("^FO50,420^FD"+ city + " " + state + " " + zipcode + " " + country + "^FS \n");
			stream.write("^CFA,15 \n");
			stream.write("^FO600,300^GB150,150,3^FS \n");
			stream.write("^FO638,340^FDPermit^FS \n");
			stream.write("^FO638,390^FD123456^FS \n");
			stream.write("^FO50,500^GB700,1,3^FS \n");
			stream.write("\n");
			stream.write("^FX Third section with QR code. \n");
			stream.write("^FO300,300^BQ,2,10^FDHM," + randomNumber + "^FS \n");
			stream.write("\n");
			stream.write("^FX Fourth section (the two boxes on the bottom). \n");
			stream.write("^FO50,900^GB700,250,3^FS \n");
			stream.write("^FO400,900^GB1,250,3^FS \n");
			stream.write("^CF0,40 \n");
			stream.write("^FO100,960^FDShipping Ctr. X34B-1^FS \n");
			stream.write("^FO100,1010^FDREF1 F00B47^FS \n");
			stream.write("^FO100,1060^FDREF2 BL4H8^FS \n");
			stream.write("^CF0,190 \n");
			stream.write("^FO485,965^FDCA^FS \n");
	  		stream.write("\n");
			stream.write("^XZ");
	  		stream.end();
			
			
	  		
			});
			var counter = 1;
			
			/* Once a match is found, the counter is set to 1 to initiate the second part of the code. 
			This ensures the second part of the code does not operate prior to the .zpl file being created*/
			
			
			
			

		};
		/* This if condition ensures that the while loop does not become infinite if someone enters
		that is not bought
		*/
		if (i>100000000000000) {
				var counter = 3;
			};
		i++; 
	};
	
	if (counter == 1){
		/* Once the file has been created, the system prints the file.
		*/
		var sys = require('sys');
		var exec = require('child_process').exec;
		function puts(error, stdout, stderr) { sys.puts(stdout) };
		exec("cat /home/sean/Desktop/QRLabel.zpl | nc 10.10.102.100 9100", puts);
		var counter = 2;

	};
	if (counter==2){
	
	/* Once the file has been printed, the user is informed what information was sent, and 
	how to go about initial troubleshooting if necessary.
	*/
	
	var clc = require('cli-color');
	console.log(clc.inverse("\n" + result.number + "\n"));
	console.log(clc.bold("	" + firstName + " " + lastName));
	console.log(clc.greenBright("	" + streetAddressLine1 + "\n" + "	" + streetAddressLine2+ "\n" + "	" + city + " " + state + " " + zipcode + " " + country ) );
	console.log(clc.blueBright("\n"+"	" +"The file was saved successfully!\n" + "	The file should have printed, if not: \n" +"	Copy ' cat /Users/Kessel/Documents/address.zpl | nc 10.10.102.100 9100 '\n"+"	" +"into the command line and hit enter. \n" ));
	var counter = 4;
	};

	
	if (counter == 3){
		var clc = require('cli-cokllor');
		console.log(clc.redBright("	 The QR code scanned does not seem to correspond to anyone in the Excel file. \n 	Check to ensure the right product was scanned."));
	};
};

});


