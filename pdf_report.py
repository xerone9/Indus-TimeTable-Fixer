from fpdf import FPDF
import datetime
import os

def generate_pdf_report(list, day, session):
    class PDF(FPDF):
        def header(self):
            date_complete = str(datetime.datetime.now()).split(".")
            date_and_time = date_complete[0].split(" ")
            date = date_and_time[0]

            # Logo
            self.image('iu_logo.png', 10, 8, 55)
            self.set_font('helvetica', 'IU', 12)
            self.cell(150)
            self.cell(10, 25, "Print Date: " + str(date), ln=0, align='L')
            self.ln(5)
            # font
            self.set_font('helvetica', 'B', 20)
            # Padding
            self.cell(70)
            # Title
            self.cell(60, 50, day + '  Classes Time Table ' + str(session), ln=0, align='C')
            self.set_font('helvetica', '', 12)
            self.cell(80, 10, '', border=False, ln=1, align='R')
            # Line break
            self.ln(15)
            # font

            self.ln(10)

            self.set_font('helvetica', 'B', 16)
            self.cell(1)

            self.set_fill_color(r=0, g=0, b=0)
            self.set_text_color(255, 255, 255)
            self.cell(30, 10, "Timings", border=True, ln=0, fill=True)
            self.cell(60, 10, "Course Title ", border=True, ln=0, fill=True)
            self.cell(60, 10, "Teacher", border=True, ln=0, fill=True)
            self.cell(40, 10, "Location", border=True, ln=1, fill=True)


    # Create a PDF object
    pdf = PDF('P', 'mm', 'A4')

    # get total page numbers
    pdf.alias_nb_pages()

    # Set auto page break
    pdf.set_auto_page_break(auto = True, margin = 50)

    #Add Page
    pdf.add_page()

    # specify font

    pdf.set_font('helvetica', 'B', 14)

    pdf.cell(1)
    pdf.set_fill_color(208, 206, 206)
    pdf.cell(190, 10, "SLOT 1", border=True, align='C', ln=1, fill=True)
    pdf.set_font('helvetica', '', 10)

    index = 0
    for row in list:
        split_values = str(row).split(" ----- ")[0]
        split_time = str(split_values).split(" - ")[0]
        index += 1
        if split_time == "12:00" or split_time == "12:15" or split_time == "12:30":
            list.insert(index - 1, "Second Slot")
            break

    index = 0
    for row in list:
        split_values = str(row).split(" ----- ")[0]
        split_time = str(split_values).split(" - ")[0]
        index += 1
        if split_time == "15:00" or split_time == "15:15" or split_time == "15:30":
            list.insert(index - 1, "Third Slot")
            break

    for item in list:
        if item == "Second Slot":
            pdf.cell(1)
            pdf.set_fill_color(208, 206, 206)
            pdf.set_font('helvetica', 'B', 14)
            pdf.cell(190, 10, "SLOT 2", border=True, align='C', ln=1, fill=True)
        elif item == "Third Slot":
            pdf.cell(1)
            pdf.set_fill_color(208, 206, 206)
            pdf.set_font('helvetica', 'B', 14)
            pdf.cell(190, 10, "SLOT 3", border=True, align='C', ln=1, fill=True)
        else:
            pdf.set_font('helvetica', '', 10)
            split_data = str(item).split(" ----- ")
            pdf.cell(1)
            pdf.cell(30, 10, str(split_data[0]), border=True, ln=0)
            if len(str(split_data[1])) > 32:
                pdf.cell(60, 10, str(split_data[1][0:32]) + " ...", border=True, ln=0)
            else:
                pdf.cell(60, 10, str(split_data[1]), border=True, ln=0)
            pdf.cell(60, 10, str(split_data[2]), border=True, ln=0)
            pdf.cell(40, 10, str(split_data[3]), border=True, ln=1)


    pdf.set_font('times', '', 12)

    desktop = os.path.expanduser("~\Desktop\\")
    saveLocation = desktop + "Scehedule For Day " + str(day) + ".pdf"
    pdf.output(saveLocation)


# list = ['09:00 - 10:00 ----- Interior Design Studio-I sssssssssssssssssssssssssssssss(Thy) ----- Sadia Perveen (Visiting) ----- Fashion theis room', '09:00 - 11:00 ----- Design Collection-I (Lab) ----- Faryal Ahsun (Fulltime) ----- Fashion thesis room', '09:00 - 11:00 ----- History of Art &amp; Culture-I ----- Zainab Abid (Visiting) ----- LEcture Room,LEcture Room', '09:00 - 11:00 ----- History of Art &amp; Culture-II ----- Afsheen Khalid (Visiting) ----- Lecture Room', '09:00 - 11:00 ----- Intro to Textile ----- Fareesa Javaid (Fulltime) ----- Lecture Room,Computer Lab', '09:00 - 12:00 ----- Business English,Business English ----- Yamna Khan (Visiting) ----- Lecture Room,LEcture Room', '09:00 - 12:00 ----- Digital Communication-II (Lab) ----- Amna Hashmi (Visiting) ----- Lecture Room,Lecture Room', '10:00 - 12:00 ----- Interior Design Studio-I (Lab) ----- Sadia Perveen (Visiting) ----- Seminar Hall', '10:00 - 13:00 ----- Collection-II (Thesis) (Lab) ----- Noor Ul Ain Shaikh (Visiting) ----- Textile Thesis Room', '11:00 - 12:00 ----- Design Collection-II (Lab) ----- Faryal Ahsun (Fulltime) ----- Textile thesis Lab', '12:30 - 14:30 ----- History of Art and Architecture-I ----- Kainat Riaz (Fulltime) ----- Thesis room,Lecture Room', '12:30 - 14:30 ----- History of Arts ----- Syeda Mona Batool Taqvi (Visiting) ----- lec room', '12:30 - 15:30 ----- Design Collection-III (Lab) ----- Faryal Ahsun (Fulltime) ----- Fashion theis room', '12:30 - 15:30 ----- English-I (Compulsory) ----- Zuraiz Akhter (Fulltime) ----- Fashion thesis room', '12:30 - 15:30 ----- History of Media Art ----- Syeda Mona Batool Taqvi (Visiting) ----- LEcture Room,LEcture Room', '12:30 - 15:30 ----- Media Laws and Ethics ----- Ayesha Naveed (Fulltime) ----- Lec room', '15:30 - 18:30 ----- Media Laws and Ethics ----- Ayesha Naveed (Fulltime) ----- Lec room', '15:30 - 18:30 ----- Media Laws and Ethics ----- Ayesha Naveed (Fulltime) ----- Lec room', '15:30 - 18:30 ----- Media Laws and Ethics ----- Ayesha Naveed (Fulltime) ----- Lec room', '15:30 - 18:30 ----- Media Laws and Ethics ----- Ayesha Naveed (Fulltime) ----- Lec room']
# generate_pdf_report(list, "Monday", "Fall 2022")