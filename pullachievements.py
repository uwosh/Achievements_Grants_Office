# generates a CSV that can be used with Word merge

import logging
log = logging.getLogger('pullachievements')

CLEANUP_CHARACTERS = False

class Award:

    achievement_name = "Faculty Development Awards"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Award.achievement_name,
            self.mydict['name-of-award'],
            self.mydict['organization-making-the-award']
            ])

    @staticmethod
    def blank():
        return "|".join([Award.achievement_name, "", ""])

    @staticmethod
    def header(n):
        return "|".join([Award.achievement_name + " " + str(n), 'name-of-award-%s' % n, 'organization-making-the-award-%s' % n])


class Book:

    achievement_name = "Publications"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Book.achievement_name,
            self.mydict['co-author-s'],
            self.mydict['editor-if-part-of-an-anthology'],
            self.mydict['book-title'],
            self.mydict['publisher'],
            self.mydict['year-published']
            ])

    @staticmethod
    def blank():
        return "|".join([Book.achievement_name, "", "", "", "", ""])

    @staticmethod
    def header(n):
        return "|".join([Book.achievement_name + " " + str(n), "co-author-s-%s" % n, "editor-if-part-of-an-anthology-%s" % n, "book-title-%s" % n, "publisher-%s" % n, "year-published-%s" % n])


class Fellowship:

    achievement_name = "Fellowships and Individual Grants"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Fellowship.achievement_name,
            self.mydict['co-awardee'],
            self.mydict['project-title'],
            self.mydict['sponsor']
            ])

    @staticmethod
    def blank():
        return "|".join([Fellowship.achievement_name, "", "", ""])
    
    @staticmethod
    def header(n):
        return "|".join([Fellowship.achievement_name + " " + str(n), "co-awardee-%s" % n, "project-title-%s" % n, "sponsor-%s" % n])


class Presentation:

    achievement_name = "Professional Presentations"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Presentation.achievement_name,
            self.mydict['co-presenters'],
            self.mydict['title-of-presentation'],
            self.mydict['name-of-conference'],
            self.mydict['city-where-presented'],
            self.mydict['state-where-presented'] if (self.mydict['state-where-presented']).lower() <> 'n/a' else "",
            self.mydict['country-where-presented'],
            "%s %s" % (self.mydict['month-presented'], self.mydict['year-presented'])
            ])

    @staticmethod
    def blank():
        return "|".join([Presentation.achievement_name, "", "", "", "", "", "", ""])

    @staticmethod
    def header(n):
        return "|".join([Presentation.achievement_name + " " + str(n), "co-presenters-%s" % n, "title-of-presentation-%s" % n, "name-of-conference-%s" % n, "city-where-presented-%s" % n, "state-where-presented-%s" % n, "country-where-presented-%s" % n, "date-presented-%s" % n])


class Exhibition:

    achievement_name = "Art Exhibitions and Performances"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Exhibition.achievement_name,
            self.mydict['co-exhibitor-s'],
            self.mydict['exhibit-performance-title'],
            self.mydict['gallery-name-or-sponsor'],
            self.mydict['city-where-presented'],
            self.mydict['state-where-presented'] if (self.mydict['state-where-presented']).lower() <> 'n/a' else "",
            self.mydict['country-where-presented'],
            "%s %s" % (self.mydict['start-month'], self.mydict['start-year'])
            ])

    @staticmethod
    def blank():
        return "|".join([Exhibition.achievement_name, "", "", "", "", "", "", ""])

    @staticmethod
    def header(n):
        return "|".join([Exhibition.achievement_name + " " + str(n), "co-exhibitor-s-%s" % n, "exhibit-performance-title-%s" % n, "gallery-name-or-sponsor-%s" % n, "city-where-presented-%s" % n, "state-where-presented-%s" % n, "country-where-presented-%s" % n, "date-presented-%s" % n])


class Appointment:

    achievement_name = "Appointments/Elections"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Appointment.achievement_name,
            self.mydict['your-title'],
            self.mydict['organization-title'],
            "%s %s" % (self.mydict['start-month'], self.mydict['start-year'])
            ])

    @staticmethod
    def blank():
        return "|".join([Appointment.achievement_name, "", "", ""])

    @staticmethod
    def header(n):
        return "|".join([Appointment.achievement_name + " " + str(n), "your-title-%s" % n, "organization-title-%s" % n, "start-date-%s" % n])


class Journal:

    achievement_name = "Journals"

    def __init__(self, adict):
        self.mydict = adict

    def __str__(self):
        return "|".join([
            Journal.achievement_name,
            self.mydict['co-author-s'],
            self.mydict['article-title'],
            self.mydict['journal-title'],
            self.mydict['month-season-of-publication'],
            self.mydict['year-of-publication'],
            self.mydict['volume-of-publication'],
            self.mydict['issue-of-publication']
            ])

    @staticmethod
    def blank():
        return "|".join([Journal.achievement_name, "", "", "", "", "", "", ""])

    @staticmethod
    def header(n):
        return "|".join([Journal.achievement_name + " " + str(n), "co-author-s-%s" % n, "article-title-%s" % n, "journal-title-%s" % n, "month-season-of-publication-%s" % n, "year-of-publication-%s" % n, "volume-of-publication-%s" % n, "issue-of-publication-%s" % n])


def pullachievements(self):

    out = []
    line = ""
    alldata_by_dept = {}
    
    NAME = 'your-name'
    
    DEPT = 'department'
    
    AWARD = 'award'
    max_awards = 0
    
    BOOK = 'book'
    max_books = 0
    
    FELLOWSHIP = 'fellowship or individual grant'
    max_fellowships = 0
    
    PRESENTATION = 'professional presentation'
    max_presentations = 0
    
    EXHIBITION = 'art exhibition or performance'
    max_exhibitions = 0
    
    APPOINTMENT = 'appointment or election'
    max_appointments = 0
    
    JOURNAL = 'journal'
    max_journals = 0
    
    awards_data = self['copy_of_awards']['awards-saved-data']
    book_data = self['book']['books-saved-data']
    fellowship_data = self['fellowships-and-individual-grants']['fellowships-and-individual-grants-saved-data']
    presentation_data = self['copy_of_professional-presentations']['professional-presentations-saved-data']
    exhibition_data = self['copy_of_art-exhibitions-and-performances']['art-exhibitions-and-performances-saved-data']
    appointment_data = self['copy_of_appointments-elections']['appointments-elections-saved-data']
    journal_data = self['journal']['journals-saved-data']

    def clean_dept(dept):
        mapping = {
            "C&I": "Curriculum and Instruction", 
            "Curriculum & Instruction" : "Curriculum and Instruction",
            "COB Accounting Dept" : "Accounting",
            "COB Management Human Resources" : "Management and Human Resources",
            "COB Mkt & Supply Chain Mgmt" : "Marketing and Supply Chain Management",
            "CON" : "College of Nursing",
            "CRIMINAL JUSTICE" : "Criminal Justice",
            "College of Educ & Human Srvcs" : "College of Education and Human Services",
            "Human Serv/Educ Leadership" : "Human Services and Educational Leadership",
            "LIBRARY - Polk" : "Polk Library",
            "Management & Human Resources" :  "Management and Human Resources",
            "Marketing & Supply Chain Management (COB)" : "Marketing and Supply Chain Management",
            "Public Admin" : "Public Administration",
            "Religious Studies & Anthro" : "Religious Studies and Anthropology",
            "Foreign Languages & Literature" : "Foreign Languages and Literature",
            "Foreign Languages & Literatures" : "Foreign Languages and Literature",
            "Human Kinetics & Health Ed" : "Human Kinetics and Health Education",
            "Human Services & Educational Leadership" : "Human Services and Educational Leadership",
            "Nursing" : "College of Nursing",
            "Physics & Astronomy" : "Physics and Astronomy",
            "Women's Studies & History" : "Women's Studies and History", 
            }
        if mapping.has_key(dept):
            return mapping[dept]
        else:
            return dept


    def output_achievement_type(line, ac_type, acs):
        if acs:
            len_acs = len(acs)
            # output the achievements the person has of this type
            for ac in acs:
                line += "|" + str(ac)
        else:
            len_acs = 0

        # fill with blanks the rest of the max len list of achievements of this type
        if ac_type == AWARD:
            for i in range(len_acs, max_awards):
                line += "|" + Award.blank()
        elif ac_type == BOOK:
            for i in range(len_acs, max_books):
                line += "|" + Book.blank()
        elif ac_type == FELLOWSHIP:
            for i in range(len_acs, max_fellowships):
                line += "|" + Fellowship.blank()
        elif ac_type == PRESENTATION:
            for i in range(len_acs, max_presentations):
                line += "|" + Presentation.blank()
        elif ac_type == EXHIBITION:
            for i in range(len_acs, max_exhibitions):
                line += "|" + Exhibition.blank()
        elif ac_type == APPOINTMENT:
            for i in range(len_acs, max_appointments):
                line += "|" + Appointment.blank()
        elif ac_type == JOURNAL:
            for i in range(len_acs, max_journals):
                line += "|" + Journal.blank()
        else:
            log.error("error handling achievement type %s row is %s line is: %s" % (ac_type, acs, line))
            raise Exception
        return line


    def removeNonAscii(s):
        return "".join(i for i in s if ord(i)<128)


    # Read all the data into a multidimensional data structure
    # organized by department, person, achievement type.  Multiple
    # achievements of the same type for a given person are not sorted
    # in any particular order.

    for row in awards_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(AWARD)):
            alldata_by_dept[dept][name][AWARD] = []
        alldata_by_dept[dept][name][AWARD].append(Award(row))
        curlen = len(alldata_by_dept[dept][name][AWARD])
        if curlen > max_awards:
            max_awards = curlen

    for row in book_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(BOOK)):
            alldata_by_dept[dept][name][BOOK] = []
        alldata_by_dept[dept][name][BOOK].append(Book(row))
        curlen = len(alldata_by_dept[dept][name][BOOK])
        if curlen > max_books:
            max_books = curlen

    for row in fellowship_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(FELLOWSHIP)):
            alldata_by_dept[dept][name][FELLOWSHIP] = []
        alldata_by_dept[dept][name][FELLOWSHIP].append(Fellowship(row))
        curlen = len(alldata_by_dept[dept][name][FELLOWSHIP])
        if curlen > max_fellowships:
            max_fellowships = curlen

    for row in presentation_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(PRESENTATION)):
            alldata_by_dept[dept][name][PRESENTATION] = []
        alldata_by_dept[dept][name][PRESENTATION].append(Presentation(row))
        curlen = len(alldata_by_dept[dept][name][PRESENTATION])
        if curlen > max_presentations:
            max_presentations = curlen

    for row in exhibition_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(EXHIBITION)):
            alldata_by_dept[dept][name][EXHIBITION] = []
        alldata_by_dept[dept][name][EXHIBITION].append(Exhibition(row))
        curlen = len(alldata_by_dept[dept][name][EXHIBITION])
        if curlen > max_exhibitions:
            max_exhibitions = curlen

    for row in appointment_data.inputAsDictionaries():
        name = row[NAME]
        dept = clean_dept(row[DEPT])
        if (not alldata_by_dept.has_key(dept)):
            alldata_by_dept[dept] = {}
        if (not alldata_by_dept[dept].has_key(name)):
            alldata_by_dept[dept][name] = {}
        if (not alldata_by_dept[dept][name].has_key(APPOINTMENT)):
            alldata_by_dept[dept][name][APPOINTMENT] = []
        alldata_by_dept[dept][name][APPOINTMENT].append(Appointment(row))
        curlen = len(alldata_by_dept[dept][name][APPOINTMENT])
        if curlen > max_appointments:
            max_appointments = curlen

    # now output a spreadsheet that can be used for Mail Merge in MS
    # Word

    # output a header line

    line = "|".join(["dept", "name"])
    for i in range(0, max_awards):
        line += "|" + Award.header(i)
    for i in range(0, max_books):
        line += "|" + Book.header(i)
    for i in range(0, max_fellowships):
        line += "|" + Fellowship.header(i)
    for i in range(0, max_presentations):
        line += "|" + Presentation.header(i)
    for i in range(0, max_exhibitions):
        line += "|" + Exhibition.header(i)
    for i in range(0, max_appointments):
        line += "|" + Appointment.header(i)
    for i in range(0, max_journals):
        line += "|" + Journal.header(i)
    out.append(line)

    # output the values

    depts = [k for k in alldata_by_dept.keys()]
    depts.sort()
    for dept in depts:
        names = [k for k in alldata_by_dept[dept].keys()]
        names.sort()
        for name in names:
            line = "|".join([dept, name])
            # handle all achievement types even if the person doesn't have all of them:
            line = output_achievement_type(line, AWARD, alldata_by_dept[dept][name][AWARD] if alldata_by_dept[dept][name].has_key(AWARD) else None)
            line = output_achievement_type(line, BOOK, alldata_by_dept[dept][name][BOOK] if alldata_by_dept[dept][name].has_key(BOOK) else None)
            line = output_achievement_type(line, FELLOWSHIP, alldata_by_dept[dept][name][FELLOWSHIP] if alldata_by_dept[dept][name].has_key(FELLOWSHIP) else None)
            line = output_achievement_type(line, PRESENTATION, alldata_by_dept[dept][name][PRESENTATION] if alldata_by_dept[dept][name].has_key(PRESENTATION) else None)
            line = output_achievement_type(line, EXHIBITION, alldata_by_dept[dept][name][EXHIBITION] if alldata_by_dept[dept][name].has_key(EXHIBITION) else None)
            line = output_achievement_type(line, APPOINTMENT, alldata_by_dept[dept][name][APPOINTMENT] if alldata_by_dept[dept][name].has_key(APPOINTMENT) else None)
            line = output_achievement_type(line, JOURNAL, alldata_by_dept[dept][name][JOURNAL] if alldata_by_dept[dept][name].has_key(JOURNAL) else None)
            if CLEANUP_CHARACTERS:
                out.append(removeNonAscii(line))
            else:
                out.append(line)            

    return "\n".join(out)

if __name__ == '__main__':
    award = Award({'department': 'Theatre',
                   'name-of-award': 'Fellowship in Playwriting',
                   'organization-making-the-award': 'Eastern University',
                   'replyto': 'someone@uwosh.edu',
                   'was-this-project-supported-by-the-faculty-development-fund': 'Yes',
                   'your-name': 'Firstname Lastname'})
    pass
