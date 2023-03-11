from GuiApplication import *
from DivWiseOverall import *
from DivWiseCurrentBacklogs import *
from DivWiseSubjectResults import *
from DivWiseSPI import *
from OverallResult import *
from CurrentBacklogs import *
from SubjectWiseResults import *
from SPI import *
from tkinter import *
#Step1:running the GUI App
obj=GuiApplication()
obj.run_app()
obj.get_subject_dict()

#Use gtu file for overall results
#Use inputfiles for div wise result
# Overall Result
obj=OverallResult()
obj.result_overall()

#Overall current backlogs
obj=CurrentBacklogs()
obj.current_backlogs_0()
obj.current_backlogs_1()
obj.current_backlogs_2()
obj.current_backlogs_3()
obj.current_backlogs_4()

#Overall SPI
obj=SPI()
obj.spi_greaterthan4()
obj.spi_greaterthan5()
obj.spi_greaterthan6()
obj.spi_greaterthan7()
obj.spi_greaterthan8()

#Overall subjects result
for subject in gb.sub_dict.values():
    obj=SubjectWiseResults(subject)
    obj.fetch_results()

#Division wise Overall Result 
obj=DivWiseOverall()#div_wise result
obj.div_result_overall()#overall result

#Division wise Current Backlogs
obj=DivWiseCurrentBacklogs()
obj.current_backlogs_0()
obj.current_backlogs_1()
obj.current_backlogs_2()
obj.current_backlogs_3()
obj.current_backlogs_4()

#Division wise SPI
obj=DivWiseSPI()
obj.spi_greaterthan4()
obj.spi_greaterthan5()
obj.spi_greaterthan6()
obj.spi_greaterthan7()
obj.spi_greaterthan8()

#Division wise results for all the subjects
print("Div-wise subject results")
for subject in gb.sub_dict.values():
    obj=DivWiseSubjectResults(subject)
    obj.fetch_results()
msg=Message(gb.root,text="Data added to csv successfully!!")
msg.config(bg='lightgreen', font=('times', 10, 'italic'))
msg.pack()
