from Tkinter import Tk
from tkFileDialog import askopenfilename
from tkFileDialog import asksaveasfilename
from tkinter import messagebox

Tk().withdraw()
messagebox.showinfo(message="Welcome to Caramelitsa's Parameter Assistant App! \n  Follow these steps: \n 1. Choose your initial file \n 2. Choose your file's name and location")

while True:
	Tk().withdraw()              # we don't want a full GUI, so keep the root window from appearing
	filename = askopenfilename() # shows an "Open" dialog box and return the path to the selected file
	try:
		ofile= open(filename,"r")
		Tk().withdraw()
		userdir=asksaveasfilename()	  #the user chooses the location and name of the final output file
		break
	except:
		break

	
OutTxt=open(userdir, "w")      #creating the final txt file

for line in ofile:
	line=line.strip()
	if line.startswith("DATAFILE"):
		line= line + "\n"
		OutTxt.write(line)
		
	#finding the position of the peak along with its error		
	pos=line.find("Pos")
	if pos!= -1:
		number=line[13:22]
		number=float(number)
		power=line[23:26]
		power=int(power)
		position=number*(10**power) 
		position=str(position) + "; "   #it converts the peak position in a string so that it can be written to the file
								
		error=line[29:33]
		error=float(error)
		powerror=line[34:]
		powerror=int(powerror)
		poserror=error*(10**powerror)
		poserror=str(poserror) + ";"   #it converts the error in a string so that it can be written to the file
							
		OutTxt.write(position)  #it writes the position to the corresponding file
		OutTxt.write(poserror) #it writes the error of the peak position to the corresponding file
					
					
							
	#finding the amplitude of the peak along with its error		
	amp=line.find("Amp")
	if amp!= -1:
		number=line[13:22]
		number=float(number)
		power=line[23:26]
		power=int(power)
		amplitude=number*(10**power) 
		amplitude=str(amplitude) + "; "   #it converts the peak amplitude in a string so that it can be written to the file
								
		error=line[29:33]
		error=float(error)
		powerror=line[34:]
		powerror=int(powerror)
		amperror=error*(10**powerror)
		amperror=str(amperror) + ";"  #it converts the error in a string so that it can be written to the file
						
		OutTxt.write(amplitude)      #it writes the amplitude to the corresponding file
		OutTxt.write(amperror) #it writes the error of the peak amplitude to the corresponding file
								
									
	#finding the width of the peak along with its error		
	wid=line.find("Wid")
	if wid!= -1:
		number=line[13:22]
		number=float(number)
		power=line[23:26]
		power=int(power)
		width=number*(10**power) 
		width=str(width) + "; "   #it converts the peak width in a string so that it can be written to the file
							
		error=line[29:33]
		error=float(error)
		powerror=line[34:]
		powerror=int(powerror)
		widerror=error*(10**powerror)
		widerror=str(widerror) + "\n"  #it converts the error in a string so that it can be written to the file
						
		OutTxt.write(width)      #it writes the width to the corresponding file
		OutTxt.write(widerror) #it writes the error of the peak width to the corresponding file		



Tk().withdraw()
messagebox.showinfo(message="Your file is ready. \n Thank you for using the app! \n Goodbye!")

OutTxt.close() #closing the final txt file

			
