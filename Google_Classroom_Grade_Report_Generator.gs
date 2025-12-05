/*

v1.2.4

Google Classroom Grade Report Generator
Works specifically for class graded by weight categories
Grade Report Constrcutor and Emailer for Google Classroom
by William Openshaw (OphaTapioka) 2023
https://github.com/WilliamOpenshaw

//last update 2025 DEC 05

*/

/*
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
*/

/*

To use this script, you'll need to replace "yourITemail@email.com" with your actual IT email address in the sendEmails function.
If you want to test sending emails to a different address, you can change it there as well.
You'll need to replace the logo image in the addLogo function with your own logo image if you want to use a different one.
This is the line where that happens: insertImageToCellWithImageBuilder(sheets[sheet], 'https://drive.google.com/uc?export=download&id=ID_OF_YOUR_LOGO_IMAGE', 1, 1);
Google appscript workspace has a 30 minute limit, so you will have to set the ListCoursesAllTabsLocalArray to run again if it stops at a specific course.
Each function can possibly take more than 30 minutes.

Your first sheet tab needs to be called "Sheet1"

In your Sheet1 tab, your student emails and guardian emails need to be in this format in these columns:

Column G                          Column H                    Column I                    Column J
Student Email	                    Guardian 1	                Guardian 2	                Guardian 3
StudentEmail1@SchoolEmail.com	    ParentEmail1@Email.com	    ParentEmail2@Email.com      ParentEmail3@Email.com
StudentEmail2@SchoolEmail.com	    ParentEmail1@Email.com	    ParentEmail2@Email.com      ParentEmail3@Email.com
StudentEmail3@SchoolEmail.com	    ParentEmail1@Email.com	    ParentEmail2@Email.com      ParentEmail3@Email.com
StudentEmail4@SchoolEmail.com	    ParentEmail1@Email.com	    ParentEmail2@Email.com      ParentEmail3@Email.com
StudentEmail5@SchoolEmail.com	    ParentEmail1@Email.com	    ParentEmail2@Email.com      ParentEmail3@Email.com


Run functions in this order:

1. ListCoursesAllTabsLocalArray
2. calculateGradesLocalArray
3. setColumns
4. addLogo
5. sendEmails

*/

// How Many Courses to Process

////////////////////////////////////////////
// CURRENTLY HAS NO EFFECT
// USES TIME LIMIT OF 25 MINUTES INSTEAD
var stopAfterNumberOfCourses            =   false             ;
var numberOfCoursesToStopAt             =   55              ;

////////////////////////////////////////////

// STARTING FROM PREVIOUS INCOMPLETE ListCourses() COURSE LOGGING

var startListCoursesFromSpecificCourse  =   false           ;
var courseNumberToContinueFrom          =   56                ;

// COURSE LOGGING AND CALCULATING GRADES ONLY ONE STUDENT
var specificStudentName                 =   "Student Name"    ;
var onlyOneStudent                      =   false             ;
//var onlyOneStudent = true;

// CALCULATE GRADES START FROM SPECIFIC STUDENT SHEET
var nameToContinueFrom                  =   "Student Name"     ;
var continueFrom                        =   false             ;
//var continueFrom = true;

// ONLY LIST COURSE NAMES IN LOG, NOTHING ELSE
var justListOfCourses                   =   false             ;

// ONLY LIST COURSE NAMES AND STUDENT NAMES
var justListOfCoursesAndNames           =   false             ;

var courseCounter = 0;

var DIALOG_TITLE = 'Example Dialog';
var SIDEBAR_TITLE = 'Example Sidebar';
var alpha = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
var assignmentName = 'no_name_set';

var courseLabel = "Course Name課程名稱:";
var assignmentNameLabel = "Assignment Name名稱:";
var weightCategoryLabel = "Weight Category分數類別:";
var weightPercentLabel = "Weight Percent % 類別比重:";
var maxPointsLabel = "Max Points滿分:";
var assignedPointsLabel = "Assigned Points成績:";

var categoryLabelEng = "Category ";
var categoryLabelChi = " 成績類別-";
var scoreWeightLabel = "Score Weight 類別比重";
var gradeInCatLabelEng = "Grade in Category out of ";
var gradeInCatLabelChi = " 此類別總成績";
var overallLabel = "Overall grade in this class 目前總成績 :";

var noGradeWeightCategoriesLabel = '沒有類別 No Grade Weight Categories';
var thisAssignmentNoCategoryLabel = '沒有類別 No Category';

var firstSheet;
var studentEmail;
var alphanumber = 1;
//first use of number for sheet position for class report
var number = 2;
//new use of number for keeping track of letter of cell
var alphanum = 0;
//new use of number for sheet position for student report
var num = 4
var cell = alpha[alphanumber] + number.toString();
var cell2 = alpha[alphanumber + 1] + number.toString();
var exists = false;
var currentStudentName;
var currentStudentSheet;
var currentStudentArray;
var sheetNames = [];
var sheets;
var textPositions = {};
var hasGradeWeightCategories = true;
var columnRange;
var alreadystudent = [];
var courseYNum = 0;
var overallYNum = 0;
var lastRowOfCourse = 0;
var categoryStartYNum = 0;
var startTime;
var thisCourseStartTime;
var endTime;
var elapsedTime;
var thisCourseElapsedTime;
var allstudents = {}
var eachstudent = ''
var categoryNum = 0;
var hasStudent = true;
var nowLastRow;
var positionSaveList;
var allStudentSheetTabsArray = {};
var numberOfActiveCourses = 0;
//var lastAssignmentRow = 0;
var optionalArgs = {
  pageSize: 60
  // Use other parameter here if needed
};
//var testList;
var emailArray;
//var currentStudentIndex = 0;
startTime = new Date();
// LIST COURSES

function ListCoursesAllTabsLocalArray()
{
  console.log('Starting Function ListCoursesAllTabsLocalArray()');

  //testList = [[],[],[]]
  //testList = [['1','2','3','4'],['5','6','7','8'],['9','10','11','12']];
  //console.log('testlist before : ' + testList)
  //testList[2][1] = '11';
  //console.log('testlist after : ' + testList)
  //return;
  /**  here pass pageSize Query parameter as argument to get maximum number of result
   * @see https://developers.google.com/classroom/reference/rest/v1/courses/list
   */
  //SpreadsheetApp.getActiveSheet().clear();
  startTime = new Date();

  if (startListCoursesFromSpecificCourse == false && justListOfCoursesAndNames == false && justListOfCourses == false)
  {
    ///////////////////////////////////////////

    // Deletes all sheets except for Sheet 1
    // You only need to do this if there might have been student name or class changes
    // or if something is not working
    //deleteAllStudentSheets()

    //////////////////////////////////////////
    makeStudentList();
    sortAllStudents();
    //console.log("Object.keys(allstudents).indexOf('StudentName') : " + Object.keys(allstudents).indexOf("StudentName"))
    //return;
    makeSheetTabForEachStudent();
    deleteEmptySheets();
  }
  else
  {
    makeStudentList();
    sortAllStudents();

    //////////////////////////////////////////////

    // enable or disable these two
    // this is to test for specific errors with a certain course
    // clears all sheets
    // If you are just continuing from a previous run, 
    // you do not need to run these two functions
    //makeSheetTabForEachStudent();
    //deleteEmptySheets();

    //////////////////////////////////////////////

    continueTextPositions();
  }
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  firstSheet = sheets[0];
  emailArray = firstSheet.getRange(1, 7, firstSheet.getLastRow(), 4).getValues();
  //console.log(allstudents)
  //return;

  // ASSIGN ALL SHEET TABS
  console.log("assigning all sheets to allStudentSheetTabsArray")
  
  if(onlyOneStudent == true)
  {
    currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(specificStudentName);
    nowLastRow =  currentStudentSheet.getLastRow();
    allStudentSheetTabsArray[specificStudentName] = currentStudentSheet.getRange(1, 1, 600, 8).getValues();
  }
  else
  {
    for (child in allstudents)
    {
      //console.log(child);
      currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(child);
      nowLastRow =  currentStudentSheet.getLastRow();
      allStudentSheetTabsArray[child] = currentStudentSheet.getRange(1, 1, 600, 8).getValues();

    }
  }
  
  //console.log("allstudents : " + allstudents[0]);
  //console.log("Jordan is student number : " + allstudents["Jordan Wang"]);
  //return;
  //console.log("allStudentSheetTabsArray['Jordan Wang'] : " + allStudentSheetTabsArray["Jordan Wang"])
  //console.log("allStudentSheetTabsArray['Jordan Wang'][0] : " + allStudentSheetTabsArray["Jordan Wang"][0])
  //console.log("allStudentSheetTabsArray : " + allStudentSheetTabsArray)
  //return;
  try
  {
    // call courses.list() method to list the courses in classroom
    const response = Classroom.Courses.list();
    const courses = response.courses;
    //console.log(courses[5].gradebookSettings.gradeCategories);
    console.log('Retrieved Course List');
    //console.log(response.courses);
    //console.log()
    //console.log('Course Work')
    //console.log()
    //const response2 = Classroom.Courses.CourseWork.list(658106163520)
    //const works = response2.courseWork
    //console.log('works : ' + works)
    if (!courses || courses.length === 0)
    {
      console.log('No courses found.');
      LabelandData('No Courses Found', '')
      return;
    }
    console.log('Number of Courses : ' + courses.length);
    console.log("Counting ACTIVE Courses");
    numberOfActiveCourses = 0;
    for (const course of courses)
    {
      if (course.courseState != 'ACTIVE')
      {
        continue;
      }
      numberOfActiveCourses++;
    }
    console.log('Number of ACTIVE Courses : ' + numberOfActiveCourses);

    //FOR EACH COURSE    
    console.log('///////////////////////////////////////////////');
    console.log('START GOING THROUGH COURSES');
    console.log('///////////////////////////////////////////////');

    // COURSES

    for (const course of courses)
    {
      thisCourseStartTime = new Date();

      if (justListOfCourses == true)
      {
        console.log('////////////////////////////////////////////////////////////////////////////////');
        console.log('Course Number : ' + courses.indexOf(course) + " out of " + courses.length);
        console.log('Course Name : ' + course.name);
        continue;
      }

      // COURSE LOGGING PROGRESS
      console.log('////////////////////////////////////////////////////////////////');
      console.log('Course Number : ' + courses.indexOf(course) + " / " + courses.length + " for all courses from 0 to " + courses.length);

      //if we want to start from a specific course
      // SET STUDENTS STARTING NEW LAST ROW POINTS
      if (startListCoursesFromSpecificCourse == true && courses.indexOf(course) < courseNumberToContinueFrom)
      {
        console.log("Skip course to start from specific course.");
        continue;
      }
      else if (startListCoursesFromSpecificCourse == true)
      {
        console.log("Set student starting points, including in progress reports");
        startListCoursesFromSpecificCourse = false;
        setTextPositionsAsLastRow();
      }

      // Continue skip if course is NOT Active
      if (course.courseState != 'ACTIVE')
      {
        console.log("Course not ACTIVE");
        continue;
      }
      // SKIPPING a specific course
      /*
      if(course.name == "2025S Life Design and Careers")
      {
        console.log("/////////////////////////////////////////////////////////////");
        console.log("/////////////////////////////////////////////////////////////");
        console.log("Skipped : " + course.name);
        console.log("/////////////////////////////////////////////////////////////");
        console.log("/////////////////////////////////////////////////////////////");
        continue;
      }
      */
      console.log('////////////////////////////////////////////////////////////////');

      courseCounter += 1;

      console.log('Number of Courses processed plus this one : ' + courseCounter);

      console.log('Course Number : ' + courses.indexOf(course) + " / " + courses.length + " for all courses from course 0 to " + courses.length);
      console.log('Course Number : ' + courses.indexOf(course) + " / " + numberOfActiveCourses + " for ACTIVE Courses Only, from course 0 to " + numberOfActiveCourses);
      
      console.log('Course Name : ' + course.name);
      //console.log('Course ID : ' + course.id);
      //console.log('Course State : ' + course.courseState);
      const workresponse = Classroom.Courses.CourseWork.list(course.id)

      if (workresponse == null || workresponse == undefined)
      {
        console.log('workresponse is null or undefiend');
        continue;
      }
      const works = workresponse.courseWork
      if (works == null || works == undefined)
      {
        console.log('works is null or undefiend');
        continue;
      }
      console.log('Retrieved Course Work List for : ' + course.name);
      console.log('Number of assignments for this course : ' + works.length);
      numberOfAssignmentsForCurrentCourse = works.length;
      ////////////////////////////////////////////////////////////////////////////////
      // GET STUDENTS LIST
      ///////////////////////////////////////////////////////////////////////////////
      const roster = []
      const options = {
        pageSize: 60
        // Use other parameter here if needed
      };

      do {
        // Get the next page of students for this course.
        var studentresponse = Classroom.Courses.Students.list(course.id, options);

        // Add this page's students to the local collection of students.
        // (Could do something else with them now, too.)
        if (studentresponse.students)
          Array.prototype.push.apply(roster, studentresponse.students);

        // Update the page for the request
        options.pageToken = studentresponse.nextPageToken;
      } while (options.pageToken);

      var students = roster;
      //console.log("studentresponse : " + studentresponse);
      console.log("studentresponse.students.length : " + studentresponse.students.length);
      //console.log("students : " + students);
      console.log("students length : " + students.length);
      console.log('Retrieved Students List for : ' + course.name);
      //IF ONLY WANT A SPECIFIC STUDENT
      if (onlyOneStudent == true)
      {
        for (student in students)
        {
          //console.log('set false');
          hasStudent = false;
          //console.log(students[student].profile.name.fullName);
          if (students[student].profile.name.fullName == specificStudentName)
          {
            hasStudent = true;
            console.log('found student; break search and continue with student : ' + students[student].profile.name.fullName);
            break;
          }
        }
        if (hasStudent == false)
        {
          console.log('student not found; continue to search next course');
          continue;
        }
      }

      if (justListOfCoursesAndNames == true)
      {
        console.log('////////////////////////////////////////////////////////////////////////////////');
        console.log('Course Number : ' + courses.indexOf(course) + " out of " + courses.length);
        console.log('Course Name : ' + course.name);
        for (student in students)
        {
          console.log(students[student].profile.name.fullName);
        }
        continue;
      }
      //console.log(0.5);
      // COUNT CATEGORIES OF COURSE
      if (!course.gradebookSettings.gradeCategories || course.gradebookSettings.gradeCategories.length === 0)
      {
        categoryNum = 0;
      }
      else
      {
        categorieslist = course.gradebookSettings.gradeCategories;
        categoryNum = Object.keys(categorieslist).length;
        //console.log('categoryNum : ' + categoryNum);
      }
      //console.log(1);
      //console.log("course.gradebookSettings.gradeCategories.length : " + course.gradebookSettings.gradeCategories.length);

      // add all available students to list

      // EACH STUDENT

      for (student in students)
      {
        
        //console.log(1.1);

        currentStudentName = students[student].profile.name.fullName;

        //IF SPECIFIC STUDENT
        if (currentStudentName != specificStudentName && onlyOneStudent == true)
        {
          continue;
        }
        console.log(currentStudentName);

        //console.log(1.2);

        // GET STUDENT SHEET FOR THIS STUDENT IN THIS COURSE
        currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentStudentName);

        //console.log(1.3);

        //Object.keys(allstudents).indexOf("Student Name")

        //console.log(1.4);

        // SET LOCAL STUDENT SHEET

        currentStudentArray = allStudentSheetTabsArray[currentStudentName];

        //console.log("allStudentSheetTabsArray[currentStudentName]: " + allStudentSheetTabsArray[currentStudentName]);

        //console.log("allStudentSheetTabsArray[currentStudentName][0]: " + allStudentSheetTabsArray[currentStudentName][0]);

        //console.log("currentStudentArray[0]: " + currentStudentArray[0]);

        //console.log("currentStudentArray: " + currentStudentArray);

        //console.log(1.5);
        //console.log("BEFORE");
        //console.log("textPositions[currentStudentName][0] : " + textPositions[currentStudentName][0]);
        //console.log("textPositions[currentStudentName][1] : " + textPositions[currentStudentName][1]);
        
        // if past beginning of sheet
        if (textPositions[currentStudentName][1] > 5)
        {
          Down(2);
        }

        // check if there is a course too near above
        // want endOfLastCourseTooNear to be false to continue, false is good
        for (var s = 2; s < 7; s++)
        {
          if (currentStudentArray[textPositions[currentStudentName][1] - s][0] != "")
          {
            //console.log("currentStudentArray[(textPositions[currentStudentName][1]) - s][0] : " + currentStudentArray[(textPositions[currentStudentName][1]) - s][0])
            //console.log("textPositions[currentStudentName][1] : " + textPositions[currentStudentName][1])
            //console.log("textPositions[currentStudentName][1] - s : " + textPositions[currentStudentName][1] - s)
            Down((5 - s) + 2);
            s = 2;
            //console.log("FOUND end of course too close, went " + ((5 - s) + 2) + " lines down")
          }
        }
        
        
        //console.log("AFTER");
        //console.log("textPositions[currentStudentName][0] : " + textPositions[currentStudentName][0]);
        //console.log("textPositions[currentStudentName][1] : " + textPositions[currentStudentName][1]);

        //console.log(1.6);
        //SetTextPositionToLeftMostColumnIfNot();
        //COURSE NAME
        // ROW and then COLUMN - currentStudentArray[ROW][COLUMN]

        // TEST
        //console.log("currentStudentArray : " + currentStudentArray);
        //console.log("textPositions[currentStudentName][0] : " + textPositions[currentStudentName][0]);
        //console.log("textPositions[currentStudentName][1] : " + textPositions[currentStudentName][1]);
        //console.log("Test for currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])]");
        //console.log("textPositions[currentStudentName].join('') : " + textPositions[currentStudentName].join(''));
        //console.log("textPositions[currentStudentName][1] - 1 : " + (parseInt(textPositions[currentStudentName][1]) - parseInt(1)).toString());
        //console.log("alpha.indexOf(textPositions[currentStudentName][0]) : " + alpha.indexOf(textPositions[currentStudentName][0]));
        //console.log("courseLabel : " + courseLabel);
        
        // Set course area START LABEL

        currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = courseLabel;

        //console.log(1.7);
        
        //color
        currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("#EDCECD");
        // SET COURSE START ROW
        courseYNum = textPositions[currentStudentName][1];
        Right(1);
        currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = course.name;
        currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setFontWeight("bold");
        //color
        currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("#EDCECD");
        //console.log(2);
        //STUDENT EMAIL
        if (alreadystudent.includes(currentStudentName) == false)
        {
          //console.log(2.1);
          //currentStudentSheet.getRange("C1").setValue('Student Email: ');
          currentStudentArray[0][2] = 'Student Email: ';
          //console.log(2.2);
          studentEmail = students[student].profile.emailAddress;
          //console.log(2.3);
          currentStudentArray[0][3] = studentEmail;
          //currentStudentSheet.getRange("D1").setValue(studentEmail);
          //console.log(2.4);
          //if (currentStudentSheet.getRange("D2").isBlank()) {
          if (currentStudentArray[1][3] == '')
          {
            //console.log(2.5);
            //for (var counter = 1; counter <= 250; counter = counter + 1) {
            for (var counter = 0; counter < emailArray.length; counter = counter + 1)
            {
              // GUARDIAN EMAIL
              //console.log(2.6);
              //if (firstSheet.getRange("G" + counter).getValue() == studentEmail) {
              if (emailArray[counter][0] == studentEmail)
              {
                //console.log(2.7);
                //currentStudentSheet.getRange("C2").setValue('Guardian Email: ');
                currentStudentArray[1][2] = 'Guardian Email: ';
                //console.log(2.8);
                //if (!firstSheet.getRange("H" + counter).isBlank()) {
                if (emailArray[counter][1] != '')
                {
                  //currentStudentSheet.getRange("D2").setValue(firstSheet.getRange("H" + counter).getValue());
                  currentStudentArray[1][3] = emailArray[counter][1];
                }
                //if (!firstSheet.getRange("I" + counter).isBlank()) {
                if (emailArray[counter][2] != '')
                {
                  //currentStudentSheet.getRange("E2").setValue(firstSheet.getRange("I" + counter).getValue());
                  currentStudentArray[1][4] = emailArray[counter][2];
                }
                //if (!firstSheet.getRange("J" + counter).isBlank()) {
                if (emailArray[counter][3] != '')
                {
                  //currentStudentSheet.getRange("F2").setValue(firstSheet.getRange("J" + counter).getValue());
                  currentStudentArray[1][5] = emailArray[counter][3];
                }
                //console.log(2.9);
                //console.log("student " + currentStudentName + "'s Guardian Email Added to their sheet");
                //console.log(2.91);
                break;
              }
            }
          }
          alreadystudent.push(currentStudentName);
        }
        //SetTextPositionToLeftMostColumnIfNot();
        //console.log(3);
        //IF NO CATEGORIES
        //console.log("course.gradebookSettings.gradeCategories : " + course.gradebookSettings.gradeCategories);
        //console.log("course.gradebookSettings.gradeCategories.length : " + course.gradebookSettings.gradeCategories.length);
        if (!course.gradebookSettings.gradeCategories || course.gradebookSettings.gradeCategories.length === 0)
        {
          Right(1);
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(noGradeWeightCategoriesLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = noGradeWeightCategoriesLabel;
          Left(2);
          Down(1);
          hasGradeWeightCategories = false;
        }
        else
        {
          Right(1);
          hasGradeWeightCategories = true;
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = "";
          Left(2);
          Down(1);
        }
        //console.log(4);
        //console.log("hasGradeWeightCategories : " + hasGradeWeightCategories);
        Down(1)

        // START OF LEFT SIDE CONTENT LABELS

        //console.log(5);
        //SetTextPositionToLeftMostColumnIfNot();
        //WRITE CATEGORIES
        if (hasGradeWeightCategories == true)
        {
          for (category in categorieslist)
          {
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(categoryLabelEng + (parseInt(category) + 1) + categoryLabelChi + (parseInt(category) + 1));
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = (categoryLabelEng + (parseInt(category) + 1) + categoryLabelChi + (parseInt(category) + 1));
            //color
            currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("#cdd9f5");
            Right(1);
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(categorieslist[category].name);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = categorieslist[category].name;
            //color
            currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("#cdd9f5");
            Down(1)
            Left(1);
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(scoreWeightLabel);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = scoreWeightLabel;
            Right(1);
            //console.log("CAT categorieslist[category].weight : " + categorieslist[category].weight)
            //console.log("CAT categorieslist[category].weight / 10000 : " + (categorieslist[category].weight/10000))
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(categorieslist[category].weight / 10000);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = categorieslist[category].weight / 10000
            Down(1)
            Left(1);
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(gradeInCatLabelEng + (categorieslist[category].weight / 10000) + gradeInCatLabelChi);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = (gradeInCatLabelEng + (categorieslist[category].weight / 10000) + gradeInCatLabelChi);
            Right(1);
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue('no input yet');
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = 'no input yet';
            Left(1);
            Down(1);
          }
          Down(1);
        }

        //SetTextPositionToLeftMostColumnIfNot();

        // OVERALL GRADE
        //console.log(6);        
        //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(overallLabel);
        currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = overallLabel;
        //color
        currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("#D8D2E7")
        Right(1);
        //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue('no input yet');
        currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = 'no input yet';
        currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setBackground("lightgreen")
        Down(2)
        Left(1)
        //console.log(6.1);
        // ASSIGNMENTS
        //console.log("numberOfAssignmentsForCurrentCourse: " + numberOfAssignmentsForCurrentCourse);
        //console.log("numberOfAssignmentsForCurrentCourse / 7: " + (numberOfAssignmentsForCurrentCourse / 7));
        //console.log("Math.ceil(numberOfAssignmentsForCurrentCourse / 7) : " + Math.ceil(numberOfAssignmentsForCurrentCourse / 7));
        for (var counter = 0; counter < Math.ceil(numberOfAssignmentsForCurrentCourse / 7); counter++)
        {
          //console.log(6.2);
          //COURSE LABELS
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(assignmentNameLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = assignmentNameLabel;
          Down(1)
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(weightCategoryLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = weightCategoryLabel;
          Down(1)
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(weightPercentLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = weightPercentLabel;
          Down(1)
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(maxPointsLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = maxPointsLabel;
          Down(1)
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(assignedPointsLabel);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = assignedPointsLabel;
          Down(2)
        }

        //console.log(6.3);
        //console.log("categorynum : " + categoryNum)
        //console.log("categoryNum * 3 : " + (categoryNum * 3))
        //console.log("5 + (categoryNum * 3) : " + (5 + (categoryNum * 3)))
        // ADJUST STUDENT TEXT POSITION FOR START OF LOGGING ASSIGNMENTS SECTION
        if (hasGradeWeightCategories == true)
        {
          Right(1)
          textPositions[currentStudentName][1] = courseYNum + 5 + (categorieslist.length) * 3;
        }
        else
        {
          Right(1)
          textPositions[currentStudentName][1] = courseYNum + 4;
        }
        //console.log(7);

        //currentStudentSheet.getRange('A1:K' + currentStudentArray.length).setValues(currentStudentArray);
        //console.log("letter of column : " + textPositions[currentStudentName][0])
        //console.log("alpha number of column : " + alpha.indexOf(textPositions[currentStudentName][0]))
        //console.log("number of columns in last row : " + currentStudentArray.length)
        //console.log("number of rows in currentarray : " + currentStudentArray[textPositions[currentStudentName][1]].length)

        // SET CHANGED LOCAL ARRAY TO GOOGLE SHEET
        //currentStudentSheet.getRange(1, 1, currentStudentArray.length, currentStudentArray[textPositions[currentStudentName][1]].length).setValues(currentStudentArray);

        //currentStudentArray = allStudentSheetTabsArray[currentStudentName];

        //console.log(" BEFORE Re-assigning");

        //console.log("currentStudentArray : " + currentStudentArray);

        //console.log("allStudentSheetTabsArray[currentStudentName] : " + allStudentSheetTabsArray[currentStudentName]);

        allStudentSheetTabsArray[currentStudentName] = currentStudentArray;

        //console.log(" AFTER Re-assigning");

        //console.log("currentStudentArray : " + currentStudentArray);

        //console.log("allStudentSheetTabsArray[currentStudentName] : " + allStudentSheetTabsArray[currentStudentName]);
        
        //return;
      }

      //console.log(8);
      //console.log('works : ' + works)

      // ASSSIGNMENTS SECTION
      console.log("Logging Assignments")
      for (assignment in works)
      {
        assignmentName = works[assignment].title;
        //console.log("assignmentName : " + assignmentName);
        const submissionsresponse = Classroom.Courses.CourseWork.StudentSubmissions.list(course.id, works[assignment].id)
        const submissions = submissionsresponse.studentSubmissions;
        //console.log(9);
        for (submission in submissions)
        {
          //console.log("Submission :" + submission + " / " + submissions.length);
          //console.log(10);
          //console.log(allstudents);
          //console.log(Object.keys(allstudents).find(key => allstudents[key] === submissions[submission].userId));
          currentStudentName = Object.keys(allstudents).find(key => allstudents[key] === submissions[submission].userId);
          //console.log("currentStudentName : " + currentStudentName);
          //console.log(11);

          //IF SPECIFIC STUDENT
          if (currentStudentName != specificStudentName && onlyOneStudent == true)
          {
            //console.log("Not the student we are looking for.");
            continue;
          }
          else if (currentStudentName == undefined)
          {
            console.log("Current Student Name is Undefined, CONTINUE skipping to next submission");
            continue;
          }

          // NEXT ROW IF THIS ROW IS FULL 
          if (alpha.indexOf(textPositions[currentStudentName][0]) > 7)
          {
            textPositions[currentStudentName][0] = 'B';
            Down(6)
          }
          // Move to the second column if text position is in first column
          if (alpha.indexOf(textPositions[currentStudentName][0]) < 1)
          {
            textPositions[currentStudentName][0] = 'B';
          }

          currentStudentArray = allStudentSheetTabsArray[currentStudentName];
          
          //console.log(12.1);
          //console.log(currentStudentName);
          //console.log(12.2);
          //console.log(textPositions);
          //console.log(12.3);
          //Assignment Name, Weight Category, Weight Percent, Max Points, Assigned Points 
          //Assignment Name
          //console.log(textPositions[currentStudentName]);
          //console.log(12.5);
          //console.log(textPositions[currentStudentName].join(''));
          //console.log(12.6);
          //console.log(assignmentName);
          //console.log(12.7);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = assignmentName;
          //lastAssignmentRow = textPositions[currentStudentName][1];
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(assignmentName);
          //console.log(12.8);
          Down(1);
          //console.log(13);
          //Weight Category, Weight Percent
          if (hasGradeWeightCategories == false)
          {
            //Weight Category
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(thisAssignmentNoCategoryLabel);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = thisAssignmentNoCategoryLabel;
            Down(1);
            //Weight Percent
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(100);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = 100;
            
            Down(1);
          }
          else if (!works[assignment].gradeCategory || works[assignment].gradeCategory === 0)
          {
            //Weight Category
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(thisAssignmentNoCategoryLabel);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = thisAssignmentNoCategoryLabel;
            Down(1);
            //Weight Percent
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(0);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = 0;
            Down(1);
          }
          else
          {
            //Weight Category
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(works[assignment].gradeCategory.name);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = works[assignment].gradeCategory.name;
            Down(1);
            //Weight Percent
            //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(works[assignment].gradeCategory.weight / 10000);
            currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = works[assignment].gradeCategory.weight / 10000;
            Down(1);
          }
          //console.log(14);
          //Max Points
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(works[assignment].maxPoints);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = works[assignment].maxPoints;
          Down(1);
          //Assigned Points
          //currentStudentSheet.getRange(textPositions[currentStudentName].join('')).setValue(submissions[submission].assignedGrade);
          currentStudentArray[textPositions[currentStudentName][1] - 1][alpha.indexOf(textPositions[currentStudentName][0])] = submissions[submission].assignedGrade;
          Right(1);
          Up(4);
          //console.log(15);
        }
      }
      //console.log(16);
      // RESET STUDENT POSITIONS AFTER EACH COURSE
      for (student in students)
      {
        //console.log(17);
        currentStudentName = students[student].profile.name.fullName;
        //IF SPECIFIC STUDENT
        if (currentStudentName != specificStudentName && onlyOneStudent == true)
        {
          continue;
        }
        currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentStudentName);
        //nowLastRow = currentStudentSheet.getLastRow();
        //column
        textPositions[currentStudentName][0] = alpha[0]
        //row
        textPositions[currentStudentName][1] += 8;
        //Down(7 + (categoryNum * 3));
        //console.log(18);
        //currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[currentStudentName]);
      }
      //categoryNum = 0;
      //console.log(19);
      // STOP COURSE ASSIGNMENT LOGGING AT CERTAIN SPOT

      endTime = new Date();
      elapsedTime = endTime - startTime;
      elapsedTime /= 60000;

      if (elapsedTime > 25 || (courseCounter >= numberOfCoursesToStopAt && stopAfterNumberOfCourses == true))
      {

        console.log('last course processed: ' + courses.indexOf(course))
        console.log('//////////////////////////////////////////////////////////////////////////////////////////////')
        console.log('25 minutes have elapsed.')
        console.log('last course name processed: ' + course.name)
        console.log('last course number processed: ' + courses.indexOf(course))
        console.log(courseCounter + ' courses processed [course counter value]');
        console.log("In Your Next Run, Please Set The Script To Run From Course Number : " + (courses.indexOf(course) + 1));
        endTime = new Date();
        elapsedTime = endTime - startTime;
        elapsedTime /= 60000;
        console.log('Assignment Inserts: Time elapsed in minutes since beginning : ' + elapsedTime);

        // ASSIGN array sheets to google sheet tabs
        console.log("assigning all sheets to allStudentSheetTabsArray")
        if (onlyOneStudent == true)
        {
          // assign one sheet
          currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(specificStudentName);
          currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[specificStudentName]);
        }
        else
        {
          // assign all sheets
          for (child in allstudents)
          {
            //console.log(child);
            currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(child);
            //nowLastRow =  currentStudentSheet.getLastRow();
            //allStudentSheetTabsArray[child] = currentStudentSheet.getRange(1, 1, 600, 8).getValues();
            currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[child]);

          }
        }

        return;
      }
      //time for each course
      endTime = new Date();
      elapsedTime = endTime - startTime;
      elapsedTime /= 60000;
      console.log('Minutes Since Start : ' + elapsedTime.toFixed(2) + " minutes");
      thisCourseElapsedTime = endTime - thisCourseStartTime;
      thisCourseElapsedTime /= 60000;

      console.log(  'Minutes THIS course took : ' +
                    thisCourseElapsedTime.toFixed(2) + " minutes");

      console.log(  'Minutes and Seconds THIS course took : ' + Math.floor(thisCourseElapsedTime) + 
                    " minutes and " + Math.floor((thisCourseElapsedTime * 60) % 60) + " seconds" );

      console.log('Seconds THIS course took : ' + Math.floor((thisCourseElapsedTime * 60)) + " seconds");
      
      console.log(  'Average course time in minutes : ' +
                    Math.floor(  elapsedTime / courseCounter) +
                    " minutes and " +
                    (((elapsedTime / courseCounter)*60)%60).toFixed(2) +
                    " seconds");
      
      console.log(  'Average course time in seconds : ' + 
                    ((elapsedTime / courseCounter) * 60).toFixed(2) + 
                    " seconds");
      
      console.log('Average course time in minutes : ' + 
                  Math.floor(elapsedTime / courseCounter) +  
                  " minutes and " +
                  (((elapsedTime / courseCounter)*60)%60).toFixed(2) +
                  " seconds");

      console.log('Total Number of Active Courses : ' + numberOfActiveCourses);

      console.log('Available time left in this run : ' + Math.floor(25 - elapsedTime) + " minutes" );

      console.log(  'Estimated time required to finish remaining courses : ' + 
                    Math.ceil((elapsedTime / courseCounter) * (numberOfActiveCourses - courseCounter)) +
                    " minutes");
      
      console.log(  'Estimated time required to finish all courses : ' + 
                    Math.ceil((elapsedTime / courseCounter) * (numberOfActiveCourses)) +
                    " minutes");

      //console.log('Estimated time to finish this run in minutes : ' + ((elapsedTime/courseCounter) * numberOfCoursesToStopAt).toFixed(2));
      //console.log('Estimated time left in this run in minutes : ' + (((elapsedTime/courseCounter) * numberOfCoursesToStopAt) - elapsedTime).toFixed(2));
      //console.log('Recommended setting for numberOfCoursesToStopAt : ' + (((30/(elapsedTime/courseCounter)).toFixed(0)) - 2));

    }
    // ASSIGN array sheets to google sheet tabs
    console.log("assigning all sheets to allStudentSheetTabsArray")
    if (onlyOneStudent == true)
    {
      // assign one sheet
      currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(specificStudentName);
      currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[specificStudentName]);
    }
    else
    {
      // assign all sheets
      for (child in allstudents)
      {
        //console.log(child);
        currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(child);
        //nowLastRow =  currentStudentSheet.getLastRow();
        //allStudentSheetTabsArray[child] = currentStudentSheet.getRange(1, 1, 600, 8).getValues();
        currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[child]);

      }
    }
    
  }
  catch (err)
  {
    // TODO (developer)- Handle Courses.list() exception from Classroom API
    // get errors like PERMISSION_DENIED/INVALID_ARGUMENT/NOT_FOUND
    console.log('CAUGHT ERROR');
    console.log('LOGGING COURSES FAILED');
    console.log('COURSE LOGGING FAILURE');
    console.log('FIX ERROR AND LOG COURSES AGAIN');
    console.log('Failed with error %s', err.message);
    console.log('STATE AT TIME OF ERROR:');
    //time for each course
      endTime = new Date();
      elapsedTime = endTime - startTime;
      elapsedTime /= 60000;
      console.log('Minutes Since Start : ' + elapsedTime.toFixed(2) + " minutes");
      thisCourseElapsedTime = endTime - thisCourseStartTime;
      thisCourseElapsedTime /= 60000;

      console.log(  'Minutes THIS course took : ' +
                    thisCourseElapsedTime.toFixed(2) + " minutes");

      console.log(  'Minutes and Seconds THIS course took : ' + Math.floor(thisCourseElapsedTime) + 
                    " minutes and " + Math.floor((thisCourseElapsedTime * 60) % 60) + " seconds" );

      console.log('Seconds THIS course took : ' + Math.floor((thisCourseElapsedTime * 60)) + " seconds");
      
      console.log(  'Average course time in minutes : ' +
                    Math.floor(  elapsedTime / courseCounter) +
                    " minutes and " +
                    (((elapsedTime / courseCounter)*60)%60).toFixed(2) +
                    " seconds");
      
      console.log(  'Average course time in seconds : ' + 
                    ((elapsedTime / courseCounter) * 60).toFixed(2) + 
                    " seconds");
      
      console.log('Average course time in minutes : ' + 
                  Math.floor(elapsedTime / courseCounter) +  
                  " minutes and " +
                  (((elapsedTime / courseCounter)*60)%60).toFixed(2) +
                  " seconds");

      console.log('Total Number of Active Courses : ' + numberOfActiveCourses);

      console.log('Available time left in this run : ' + Math.floor(25 - elapsedTime) + " minutes" );

      console.log(  'Estimated time required to finish remaining courses : ' + 
                    Math.ceil((elapsedTime / courseCounter) * (numberOfActiveCourses - courseCounter)) +
                    " minutes");
      
      console.log(  'Estimated time required to finish all courses : ' + 
                    Math.ceil((elapsedTime / courseCounter) * (numberOfActiveCourses)) +
                    " minutes");
      console.log("currentStudentName : " + currentStudentName);
      console.log("textPositions[currentStudentName][0] : " + textPositions[currentStudentName][0]);
      console.log("textPositions[currentStudentName][1] : " + textPositions[currentStudentName][1]);

      // ASSIGN array sheets to google sheet tabs
      console.log("assigning all sheets to allStudentSheetTabsArray")
      if (onlyOneStudent == true)
      {
        // assign one sheet
        currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(specificStudentName);
        currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[specificStudentName]);
      }
      else
      {
        // assign all sheets
        for (child in allstudents)
        {
          //console.log(child);
          currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(child);
          //nowLastRow =  currentStudentSheet.getLastRow();
          //allStudentSheetTabsArray[child] = currentStudentSheet.getRange(1, 1, 600, 8).getValues();
          currentStudentSheet.getRange(1, 1, 600, 8).setValues(allStudentSheetTabsArray[child]);

        }
      }

  }
  endTime = new Date();
  elapsedTime = endTime - startTime;
  elapsedTime /= 60000;
  console.log('LISTING AND LOGGING COURSES PROCESS HALTED OR COMPLETE');
  console.log('Assignment Inserts: Time elapsed in minutes since beginning : ' + elapsedTime);

  return;

  // CALCULATE GRADES
  calculateGradesLocalArray();
  //calculateGrades();

  endTime = new Date();
  elapsedTime = endTime - startTime;
  elapsedTime /= 60000;
  console.log('Calculate Grades: Time elapsed in minutes since beginning : ' + elapsedTime);

  // SET COLUMNS
  setColumns();

  endTime = new Date();
  elapsedTime = endTime - startTime;
  elapsedTime /= 60000;
  console.log('After Setting Column Width: Time elapsed in minutes since beginning : ' + elapsedTime);

  // ADD LOGO
  addLogo();
}

function calculateGradesLocalArray()
{
  console.log('Starting Function calculateGradesLocalArray()');
  startTime = new Date();
  console.log('Calculate Grades Local Array Start');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  totalNumberOfSheets = sheets.length;
  for (sheet in sheets)
  {
    currentSheetName = sheets[sheet].getName();

    //IF SPECIFIC STUDENT
    if (currentSheetName == 'Sheet1' || (onlyOneStudent == true && currentSheetName != specificStudentName))
    {
      console.log('skip : ' + currentSheetName);
      continue;
    }
    courseYNum = 3;
    overallYNum = 5;

    //FOR CONTINUING FROM A POINT

    if (currentSheetName == nameToContinueFrom && continueFrom == true)
    {
      pickupStudentArrived = true;
    }

    if (pickupStudentArrived == false && continueFrom == true)
    {
      continue;
    }

    // MAKE LOCAL LIST

    nowLastRow = sheets[sheet].getLastRow();
    //console.log("nowLastRow : " + nowLastRow);

    if (sheets[sheet].getRange(nowLastRow, 1).getValue() != "END")
    {
      sheets[sheet].getRange(nowLastRow + 3, 1).setValue("END");
    }

    nowLastRow = sheets[sheet].getLastRow();
    //console.log("nowLastRow : " + nowLastRow)

    numbersList = sheets[sheet].getRange('A1:H' + nowLastRow).getValues();
    //console.log(numbersList)

    numbersList[0][1] = currentSheetName;
    console.log('sheet number : ' + parseInt(sheet) + "/" + totalNumberOfSheets);
    currentPointsRow = 0;

    reachedEnd = false;

    while (reachedEnd == false)
    {

      //currentPointsRow = 40;
      //console.log('sheet name : ' + currentSheetName);
      //console.log('sheet number : ' + parseInt(sheet));
      //console.log('last row number : ' + nowLastRow);
      for (let i = currentPointsRow + 1; i < nowLastRow; i++)
      {
        //console.log('////////////////////////////////////////////////');
        //console.log('check for assignedpoints row : ' + i);
        //console.log('i : ' + i);
        //console.log('row cell value : ' + numbersList[i][0]);

        if (i == nowLastRow - 1)
        {
          //console.log('reached end');
          reachedEnd = true;
          break;
        }

        //if (sheets[sheet].getRange(i, 1).getValue() == assignedPointsLabel && !sheets[sheet].getRange(i, 1).isBlank())
        // IF ASSIGNED POINTS ROW IS FOUND
        if (numbersList[i][0] == assignedPointsLabel)
        {
          for (let c = i; c > 1; c--)
          {
            if (numbersList[c][0] == courseLabel)
            {
              courseYNum = c;
              //console.log("courseYNum set to : " + c);
              break;
            }
          }
          //console.log("courseYNum : " + courseYNum);
          currentPointsRow = i;
          //console.log("numbersList[courseYNum][1] : " + numbersList[courseYNum][1])
          //console.log("numbersList[courseYNum][2] : " + numbersList[courseYNum][2])
          if (numbersList[courseYNum][2] == noGradeWeightCategoriesLabel)
          {
            noCategory = true;
          }
          else
          {
            noCategory = false;
          }
          //console.log('currentPointsRow : ' + currentPointsRow);
          //console.log("noCategory : " + noCategory);
          break;
        }

      }

      // CHECK IF WE HAVE REACHED THE END
      if (reachedEnd == true)
      {
        //console.log("reachedEnd is true, break")
        break;
      }

      // ITERATE THROUGH ASSIGNMENT SCORE ROW
      if (noCategory == false)
      {
        //console.log('class has categories');
        for (let j = 1; j < 8; j++)
        {
          //console.log("j : " + j);
          //console.log('checking row : ' + currentPointsRow + ' , and column : ' + j);
          // IF ASSIGNED POINTS ROW CURRENT COLUMN CELL POINTS IS BLANK OR NO CATEGORY
          if (
            numbersList[currentPointsRow][j] === '' ||
            numbersList[currentPointsRow - 1][j] === '' ||
            numbersList[currentPointsRow - 3][j] === thisAssignmentNoCategoryLabel
          )
          {
            //console.log('numbersList[currentPointsRow][j] : ' + numbersList[currentPointsRow][j])
            //console.log('numbersList[currentPointsRow - 1][j] : ' + numbersList[currentPointsRow - 1][j])
            //console.log('numbersList[currentPointsRow - 3][j] : ' + numbersList[currentPointsRow - 3][j])
            //console.log('blank so continue, class has categories')
            continue;
          }
          // IF ENTRY POINT IS ALL BLANK - END OF ASSIGNMENTS IN THIS ROW
          if (
            numbersList[currentPointsRow][j] === '' &&
            numbersList[currentPointsRow - 2][j] === '' &&
            numbersList[currentPointsRow - 3][j] === '' &&
            numbersList[currentPointsRow - 4][j] === ''
          )
          {
            //console.log('blank entry so skip whole row')
            break;
          }
          // IF NOT BLANK THEN RECORD SCORE
          currentcategory = numbersList[currentPointsRow - 3][j];
          //console.log('currentcategory : ' + currentcategory);
          currentMaxPoints = numbersList[currentPointsRow - 1][j];
          //console.log('currentMaxPoints : ' + currentMaxPoints);
          currentAssignedPoints = numbersList[currentPointsRow][j];
          //console.log('currentAssignedPoints : ' + currentAssignedPoints);
          //console.log('variables assigned');

          //Find Place in CATEGORIES to insert score
          for (let c = currentPointsRow; c > 4; c--)
          {
            //console.log('Looking for place to put score at row : ' + c)
            //if (sheets[sheet].getRange(c, 2).getValue() == currentcategory && sheets[sheet].getRange(c, 1).getValue() != weightCategoryLabel)
            if (numbersList[c][1] == currentcategory && numbersList[c][0] != weightCategoryLabel)
            {
              //Assigned Points
              //if (sheets[sheet].getRange(c + 2, 3).isBlank())
              if (numbersList[c + 2][2] == '')
              {
                numbersList[c + 2][2] = currentAssignedPoints;
                //console.log('assigned points to no input yet in cell');
                //console.log('column : ' + 3);
                //console.log('row : ' + (c + 2));
                //console.log('assigned value changed from empty to : ' + currentAssignedPoints);
              }
              else
              {
                lastNum = numbersList[c + 2][2]
                numbersList[c + 2][2] = lastNum + currentAssignedPoints;
                //console.log('added points to cell');
                //console.log('column : ' + 3);
                //console.log('row : ' + (c + 2));
                //console.log('assigned value changed from : ' + lastNum + ' to ' + (lastNum + currentAssignedPoints));
              }
              //Max Points
              //if (sheets[sheet].getRange(c + 2, 4).isBlank())
              if (numbersList[c + 2][3] === '')
              {
                numbersList[c + 2][3] = currentMaxPoints;
                //console.log('no input yet set to max points in cell');
                //console.log('column : ' + 4);
                //console.log('row : ' + (c + 2));
                //console.log('max value changed from empty to : ' + currentMaxPoints);
              }
              else
              {
                lastNum = numbersList[c + 2][3];
                numbersList[c + 2][3] = lastNum + currentMaxPoints;
                //console.log('added max points to cell');
                //console.log('column : ' + 4);
                //console.log('row : ' + (c + 2));
                //console.log('max value changed from : ' + lastNum + ' to ' + (lastNum + currentMaxPoints));
              }
              //console.log('Assigned Points Current Total : ' + numbersList[c + 2][3])
              //console.log('Max Points Current Total : ' + numbersList[c + 2][3])
              //console.log("////////////////////////////////////////////////////////")
              break;
            }
            else if (numbersList[c][0] === courseLabel)
            {
              //console.log('hit next course, continue next loop');
              break;
            }
            else
            {
              //console.log('Did not deposit score');
            }
          }
        }
      }
      // IF COURSE HAS NO WEIGHT CATEGORIES
      else if (noCategory == true)
      {
        ////console.log('class does not have categories');
        for (let j = 1; j < 8; j++)
        {
          ////console.log('score column : ' + j + '/' + nowLastColumn);
          /*
          if (j >= nowLastColumn) {
            //console.log('last column; break');
            continue;
          }
          */
          // IF THERE IS NO ASSIGNED OR MAX SCORE - CANT PROCESS THIS SCORE - NEXT ASSIGNMENT
          // IF ASSIGNED POINTS ROW CURRENT COLUMN CELL POINTS IS BLANK OR MAX POINTS BLANK
          //console.log()
          if (
            numbersList[currentPointsRow][j] === '' ||
            numbersList[currentPointsRow - 1][j] === ''
          )
          {
            //console.log('blank so continue, no categories for this class')
            continue;
          }
          // IF ENTRY POINT IS ALL BLANK - END OF ASSIGNMENTS IN THIS ROW
          if (
            numbersList[currentPointsRow][j] == '' &&
            numbersList[currentPointsRow - 2][j] == '' &&
            numbersList[currentPointsRow - 3][j] == '' &&
            numbersList[currentPointsRow - 4][j] == ''
          )
          {
            ////console.log('blank entry so skip whole row')
            break;
          }
          //console.log("Assignment Name : " + numbersList[currentPointsRow - 4][j])
          // IF NOT BLANK THEN RECORD SCORE
          currentcategory = numbersList[currentPointsRow - 3][j];
          //console.log('currentcategory : ' + currentcategory);
          currentMaxPoints = numbersList[currentPointsRow - 1][j];
          //console.log('currentMaxPoints : ' + currentMaxPoints);
          currentAssignedPoints = numbersList[currentPointsRow][j];
          //console.log('currentAssignedPoints : ' + currentAssignedPoints);
          //console.log('variables assigned');

          //Assigned Points
          //if (sheets[sheet].getRange(currentPointsRow + 1, 3).isBlank())
          if (numbersList[courseYNum + 2][2] == '')
          {
            numbersList[courseYNum + 2][2] = currentAssignedPoints;
            //console.log('assigned points to no input yet in cell : (' + (c + 2) + ', 3)');
          }
          else
          {
            lastNum = numbersList[courseYNum + 2][2];
            numbersList[courseYNum + 2][2] = lastNum + currentAssignedPoints;
            //console.log('added points to cell : (' + (c + 2) + ', 3)');
          }
          //Max Points
          //if (sheets[sheet].getRange(currentPointsRow + 1, 4).isBlank())
          if (numbersList[courseYNum + 2][3] == '')
          {
            numbersList[courseYNum + 2][3] = currentMaxPoints;
            //console.log('no input yet set to max points in cell : (' + (c + 2) + ', 4)');
          }
          else
          {
            lastNum = numbersList[courseYNum + 2][3];
            numbersList[courseYNum + 2][3] = lastNum + currentMaxPoints;
            //console.log('added max points to cell : (' + (c + 2) + ', 4)');
          }
        }
      }
      if (reachedEnd == true)
      {
        break;
      }
    }
    // PART 2
    // TALLY TOTALS
    overallMax = 100;
    scoreChanged = false;
    console.log('Tallying totals for : ' + currentSheetName)
    ////console.log('numbersList[231][1] : ' + numbersList[231][1])

    for (let p = 3; p < nowLastRow; p++)
    {
      /*
      if (numbersList[p][1] == 'no input yet')
      {
        console.log('looking for no input yet, checking row : ' + p);
        console.log('numbersList[p][1] : ' + numbersList[p][1]);
        console.log('numbersList[p][2] : ' + numbersList[p][2]);
        console.log('numbersList[p][3] : ' + numbersList[p][3]);
        console.log('numbersList[p -1 ][1] : ' + numbersList[p - 1][1]);
        console.log('numbersList[p][0] : ' + numbersList[p][0]);
      }
      */
      // If there are assigned and max points for this category
      if (numbersList[p][1] === 'no input yet' &&
        numbersList[p][2] !== '' &&
        numbersList[p][3] !== '' &&
        numbersList[p - 1][1] !== '' &&
        numbersList[p - 1][1] !== "#NUM!" &&
        numbersList[p][0] !== overallLabel
      )
      {
        allassignedpoints = numbersList[p][2];
        allmaxpoints = numbersList[p][3];
        thisweight = numbersList[p - 1][1]
        thispoints = (allassignedpoints / allmaxpoints) * thisweight
        numbersList[p][1] = thispoints.toFixed(2)
        overallcounter += thispoints;
        // clear tallies
        ////console.log('numbersList[p][2]');
        ////console.log(numbersList[p][2]);
        numbersList[p][2] = '';
        ////console.log('numbersList[p][3]');
        ////console.log(numbersList[p][3]);
        numbersList[p][3] = '';
        //console.log('Category overall score set');
        scoreChanged = true;
      }
      // THIS GRADE CATEGORY HAS NO ASSIGNED GRADES 
      else if (
        numbersList[p][1] === 'no input yet' &&
        numbersList[p][2] === '' &&
        numbersList[p][3] === '' &&
        numbersList[p - 1][1] !== '' &&
        numbersList[p - 1][1] !== "#NUM!" &&
        numbersList[p][0] !== overallLabel
      )
      {
        //thispoints = sheets[sheet].getRange(p - 1, 2).getValue();
        //overallcounter += thispoints;
        //subtract this grade from total
        overallMax -= numbersList[p - 1][1];
        numbersList[p][1] = '這個成績類別沒有功課有收到成績 No Assigned Grades in this Category Yet';
        //console.log('No Assigned Grades in this category, so it doesnt count towards total');
      }
      // OVERALL GRADE OF COURSE
      else if (numbersList[p][0] === overallLabel)
      {
        //console.log("//////////////////////////////////////////////")
        //console.log("REACHED OVERALL GRADE AT ROW : " + p)
        //console.log("//////////////////////////////////////////////")
        overallYNum = p;
        //find out where last course name is so we can highlight cells
        for (let i = overallYNum; i > 2; i--)
        {
          if (numbersList[i][0] == courseLabel)
          {
            courseYNum = i;
            //console.log("courseYNum set to : " + i);
            break;
          }
        }

        for (let i = overallYNum; i < nowLastRow - 1; i++)
        {
          if (i + 1 >= nowLastRow || i + 2 >= nowLastRow || i + 3 >= nowLastRow)
          {
            lastRowOfCourse = nowLastRow;
          }
          else if (numbersList[i][0] == assignedPointsLabel && numbersList[i + 2][0] == "")
          {
            lastRowOfCourse = i;
            //console.log("lastRowOfCourse : " + i);
            break;
          }
        }

        if (lastRowOfCourse <= overallYNum)
        {
          for (let i = overallYNum; i < nowLastRow - 1; i++)
          {
            if (i + 1 >= nowLastRow || i + 2 >= nowLastRow || i + 3 >= nowLastRow)
            {
              lastRowOfCourse = nowLastRow;
            }
            else if (numbersList[i][0] == assignedPointsLabel && numbersList[i + 2][0] == "")
            {
              lastRowOfCourse = i;
              //console.log("lastRowOfCourse : " + i);
              break;
            }
          }
        }
        //console.log("course Y Num : " + courseYNum)
        //console.log("overall Y Num : " + overallYNum)

        categoryStartYNum = courseYNum + 2;

        //highlight cells
        //set category section border
        //set overall section border

        //console.log("courseYNum : " + courseYNum);
        //console.log("overallYNum : " + overallYNum);
        //console.log("lastRowOfCourse : " + lastRowOfCourse);
        //console.log("categoryStartYNum : " + categoryStartYNum);
        //console.log("lastRowOfCourse - (overallYNum + 1) : " + (lastRowOfCourse - (overallYNum + 1)));
        //console.log("nowlastrow : " + nowLastRow);

        if (lastRowOfCourse - courseYNum > 1)
        {
          //sheets[sheet].getRange(courseYNum + 2, 1, categoryStartYNum - (courseYNum + 1), 1).setBackground('lightyellow');
          sheets[sheet].getRange(overallYNum + 3, 1, lastRowOfCourse - (overallYNum + 1), 1).setBackground('lightyellow');
          // SET BORDER
          // setBorder(top, left, bottom, right, vertical, horizontal, color, style)
          // double line border around whole course
          sheets[sheet].getRange(courseYNum + 1, 1, (lastRowOfCourse - courseYNum) + 1, 8).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.DOUBLE);
          // thick border for OVERALL SCORE
          sheets[sheet].getRange(overallYNum + 1, 1, 1, 2).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
          // Bottom border under assignment area
          sheets[sheet].getRange(overallYNum + 2, 1, 1, 8).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        }

        if ((overallYNum - (categoryStartYNum + 2)) > 0)
        {
          // thick border around category section
          sheets[sheet].getRange(categoryStartYNum + 1, 1, overallYNum - (categoryStartYNum + 1), 2).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        }
        // OVERALL GRADE
        // IF this class HAS CATEGORIES
        // space after overall grade will be blank, because if it has no categories, the assigned points and max points will be in the cells one over left and two over left from the overall score value cell.
        //console.log('numbersList[p]' + numbersList[p])
        //console.log('numbersList[p][2] : ' + numbersList[p][2])
        if (numbersList[p][2] === '')
        {
          //console.log("//////////////////////////////////////////////")
          //console.log("SETTING CATEGORY OVERALL GRADE : " + p)
          //console.log("//////////////////////////////////////////////")
          // 100
          // IF OVERALLCOUNTER is zero and STUDENT HAS NO ASSIGNED GRADES, SO FULL MARKS
          //console.log('scoreChanged : ' + scoreChanged);
          if (overallcounter == 0 && scoreChanged == false)
          {
            finalScore = 100;
            numbersList[p][1] = 100;
            //console.log('Overall score tally is zero, and no marks, so overall score set to 100');
            //console.log('unused overall counter is zero');
          }
          // ZERO
          // IF OVERALLCOUNTER is zero, STUDENT HAS ASSIGNED GRADES, WHICH ARE ZERO, SO Zero
          else if (overallcounter == 0 && scoreChanged == true)
          {
            finalScore = 0;
            numbersList[p][1] = 0;
            //console.log('Overall score tally is zero, but there are grades, overall score set to 0, because grade is zero');
            //console.log('overall counter is zero');
          }
          // NORMAL OVERALL SCORE GRADE
          // SET CAULCULATED STUDENT SCORE
          // IF OVERALLCOUNTER IS NOT ZERO, STUDENT HAS ASSIGNED GRADES
          else
          {
            finalScore = 100 * (overallcounter / overallMax);
            numbersList[p][1] = finalScore.toFixed(2);
            //console.log('Overall Tally is not blank, used overall counter');
            //console.log('overall counter : ' + overallcounter);
            //console.log('overall max : ' + overallMax);
            //console.log('Final Score : ' + finalScore);
            //console.log('overallcounter/overallMax : ' + (overallcounter / overallMax));
          }

          //highlight overall grade color
          if (finalScore >= 60)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("lightgreen")
          }
          else if (finalScore >= 50 && finalScore < 60)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("orange")
          }
          else if (finalScore < 50)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("red")
          }
          // clear tallies 
          numbersList[p][2] = '';
          numbersList[p][3] = '';
          overallcounter = 0;
          //console.log('Category Class Overall Score set');
          overallMax = 100;
          scoreChanged = false;
        }
        // IF NO CATEGORIES AND SCORE IS TOTAL
        else if (numbersList[p][2] !== '')
        {
          //console.log("p : " + p)
          allassignedpoints = numbersList[p][2];
          //console.log("allassignedpoints : " + allassignedpoints)
          allmaxpoints = numbersList[p][3];
          //console.log("allassignedpoints : " + allmaxpoints)
          thispoints = (allassignedpoints / allmaxpoints) * 100;
          //console.log("thispoints : " + thispoints)
          if (thispoints == 0)
          {
            numbersList[p][2] = 100;
          }
          else
          {
            numbersList[p][1] = thispoints.toFixed(2);
          }
          //highlight total grade color
          if (thispoints > 60)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("lightgreen")
          }
          else if (thispoints > 50 && thispoints < 60)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("orange")
          }
          else if (thispoints < 50)
          {
            sheets[sheet].getRange(p + 1, 2).setBackground("red")
          }
          // clear tallies 
          numbersList[p][2] = '';
          numbersList[p][3] = '';
          //console.log('Total overall score set');
        }
      }
    }
    // clear student ID
    numbersList[1][0] = '';
    numbersList[1][1] = '';
    nowLastRow = sheets[sheet].getLastRow();
    //console.log("nowLastRow : " + nowLastRow);
    //console.log("numbersList.length : " + numbersList.length);
    //console.log('sheetnumber : ' + sheet);
    ////console.log('numbersList : ' + numbersList)

    sheets[sheet].getRange('A1:H' + nowLastRow).setValues(numbersList);
    //sheets[sheet].getRange('A:H').setValues(numbersList);
    elapsedTime = endTime - startTime;
    elapsedTime /= 60000;
    if (elapsedTime > 25)
    {
      console.log("25 minutes elapsed")
      console.log("Last Completed Sheet : " + currentSheetName)
      return;
    }

  }
}

function setColumns()
{
  console.log('Starting Function setColumns()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() == 'Sheet1')
    {
      continue;
    }
    //IF SPECIFIC STUDENT
    if (sheets[sheet].getName() != specificStudentName && onlyOneStudent == true)
    {
      continue;
    }
    nowLastColumn = sheets[sheet].getLastColumn()
    console.log("adjusting column width and wrapping for sheet : " + sheets[sheet].getName());
    sheets[sheet].setColumnWidths(1, nowLastColumn, 167);
    columnRange = sheets[sheet].getRange(1, 1, sheets[sheet].getLastRow(), nowLastColumn);
    columnRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheets[sheet].autoResizeRows(2, 10);
    sheets[sheet].getRange("A:H").setNumberFormat("General");
  }
  console.log('Columns set');
}

function addLogo()
{
  console.log('Starting Function addLogo()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {

    currentStudentName = sheets[sheet].getName();

    if (continueFrom == true && currentStudentName != nameToContinueFrom)
    {
      continue;
    }
    else if (continueFrom == true && currentStudentName == nameToContinueFrom)
    {
      continueFrom = false;
    }

    if (currentStudentName == "Sheet1")
    {
      continue;
    }
    if (sheets[sheet].getName() != specificStudentName && onlyOneStudent == true)
    {
      continue;
    }
    sheets[sheet].insertRowBefore(1);
    sheets[sheet].getRange("A1:D1").mergeAcross();
    sheets[sheet].setRowHeight(1, 100);
    insertImageToCellWithImageBuilder(sheets[sheet], 'https://drive.google.com/uc?export=download&id=ID_OF_YOUR_LOGO_IMAGE', 1, 1);
    sheets[sheet].autoResizeRows(2, 10);
    console.log('Added logo for sheet : ' + sheet + ' out of ' + sheets.length);
  }

}

var email = 'yourITemail@email.com';
var emailSet = false;
var emailStudentFound = false;
var emailStudentToContinueFrom = 'not set here';

function SendEmails()
{
  console.log('Starting Function SendEmails()');
  var currentDate = new Date();
  var previousMonth = currentDate.getMonth() - 1;
  //var previousMonth = - 1 ;
  if (previousMonth == -1)
  {
    previousMonth = 11;
  }
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  emailStudentFound = false;
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {

    currentStudentName = sheets[sheet].getName();
    emailStudentToContinueFrom = 'StudentName';
    //console.log(1);
    if (currentStudentName == "Sheet1")
    {
      continue;
    }
    //console.log(2);
    /*
    //IF ONLY TRYING TO DO ONE
    if (currentStudentName != "Alexander Cheng")
    {
      continue;
    }
    */

    // CONTINUE FROM POINT
    /*
    if (currentStudentName != emailStudentToContinueFrom && emailStudentFound == false)
    {
      continue;
    }
    else if (currentStudentName == emailStudentToContinueFrom)
    {
      emailStudentFound = true;
    }
    */

    //console.log(3);
    console.log('Sending Email for sheet : ' + sheet + ' out of ' + sheets.length);

    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var ssID = ss.getId();
    //console.log("ssid : " + ssID);
    var sheetgId = sheets[sheet].getSheetId();
    //console.log("sheetgId : " + sheetgId);
    var sheetName = sheets[sheet].getName();
    //console.log("sheetName : " + sheetName);
    var token = ScriptApp.getOAuthToken();
    //var email = 'yourITemail@email.com';
    //var email = 
    var subject = currentStudentName + "'s Score Report for " + months[previousMonth];
    var body = "To a Guardian of " + currentStudentName + ", your child's score report is attached to this email as a PDF.";
    

    
    // CALCULATE PAGE HEIGHT SECTION
    /*
    // Get dynamic range
    var lastRow = sheets[sheet].getLastRow();
    
    // Calculate dynamic height in inches (A4 content area + data height)
    var topMargin = 0.2;
    var bottomMargin = 0.2;
    var a4Width = 8.27;  // A4 width in inches (fixed)
    //var a4Height = 11.69;  // Standard A4 height in inches

    //var pageHeight = topMargin + totalInches + bottomMargin + 0.1;  // +0.1 buffer for safety
    var pageHeight = 0;  // +0.1 buffer for safety

    for (var row = 1; row <= lastRow; row++)
    {
      //console.log(row)
      pageHeight += sheets[sheet].getRowHeight(row);
      //console.log(pageHeight)
    }
    console.log(row)
    console.log(pageHeight)
    pageHeight /= 120;
    pageHeight += topMargin + bottomMargin + 0.1;
    console.log(pageHeight)  

    */
    

    var url = "https://docs.google.com/spreadsheets/d/" +
      ssID +
      "/export?" +
      "format=xlsx" +
      "&gid=" + sheetgId +
      // TO CONTROL PAGE HEIGHT
      //"&size=8.5x" + pageHeight.toFixed(1) +
      "&size=8.5x60.0" +
      "&portrait=true" +
      '&top_margin=0.2' +
      '&bottom_margin=0.2' +
      '&left_margin=0.2' +
      '&right_margin=0.2' +
      "&exportFormat=pdf";
      
    //console.log("url : " + url);
    var result = UrlFetchApp.fetch(url,
    {
      headers:
      {
        'Authorization': 'Bearer ' + token
      }
    });
    var contents = result.getContent();

    for (z = 0; z < 3; z++)
    {
      emailSet = false;
      if (!sheets[sheet].getRange("D2").isBlank() && z == 0)
      {
        subject = "Your " + months[previousMonth] + " Grade Report";
        body = " Hi " + currentStudentName + ", your " + months[previousMonth] + " grade report is attached to this email as a PDF."
        email = sheets[sheet].getRange("D2").getValue();
        emailSet = true;
      }
      else if (!sheets[sheet].getRange("D3").isBlank() && z == 1)
      {
        subject = currentStudentName + "'s VIS Grade Report for " + months[previousMonth];
        body = "To a Guardian of " + currentStudentName + ", your child's grade report is attached to this email as a PDF.";
        email = sheets[sheet].getRange("D3").getValue();
        emailSet = true;
      }
      else if (!sheets[sheet].getRange("E3").isBlank() && z == 2)
      {
        subject = currentStudentName + "'s VIS Grade Report for " + months[previousMonth];
        body = "To a Guardian of " + currentStudentName + ", your child's grade report is attached to this email as a PDF.";
        email = sheets[sheet].getRange("E3").getValue();
        emailSet = true;
      }
      else if (!sheets[sheet].getRange("F3").isBlank() && z == 3)
      {
        subject = currentStudentName + "'s VIS Grade Report for " + months[previousMonth];
        body = "To a Guardian of " + currentStudentName + ", your child's grade report is attached to this email as a PDF.";
        email = sheets[sheet].getRange("F3").getValue();
        emailSet = true;
      }
      email = email.replaceAll(" ", "");
      email = email.trim();

      //if (emailSet == true && email == 'emailAddress@email.com')
      if (emailSet == true)
      {
        //ENABLE for testing self-sending
        //email = yourITemail@email.com;
        MailApp.sendEmail(email, subject, body,
        {
          attachments: [
          {
            fileName: sheetName + ".pdf",
            content: contents,
            mimeType: "application//pdf"
          }]
        });
        console.log(sheetName + ' Grade Report Email sent to : ' + email);
        Utilities.sleep(3000);
      }
    }
  }
}
// this function is incomplete and probably will not be used
function estimateEmailSheetLength()
{
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {

    currentStudentName = sheets[sheet].getName();
    emailStudentToContinueFrom = 'StudentName';
    //console.log(1);
    if (currentStudentName == "Sheet1")
    {
      continue;
    }

    // Get dynamic range
    var lastRow = sheets[sheet].getLastRow();
    
    // Calculate dynamic height in inches (A4 content area + data height)
    var topMargin = 0.2;
    var bottomMargin = 0.2;
    var a4Width = 8.27;  // A4 width in inches (fixed)
    //var a4Height = 11.69;  // Standard A4 height in inches

    //var pageHeight = topMargin + totalInches + bottomMargin + 0.1;  // +0.1 buffer for safety
    var pageHeight = 0;  // +0.1 buffer for safety

    for (var row = 1; row <= lastRow; row++)
    {
      //console.log(row)
      pageHeight += sheets[sheet].getRowHeight(row);
      //console.log(pageHeight)
    }
    console.log(row)
    console.log(pageHeight)
    pageHeight /= 120;
    pageHeight += topMargin + bottomMargin + 0.1;
    console.log(pageHeight)

    break;
  }
}

var sheetLocalArray;
var sendThisReport;

function sendFailingStudentReportsToEthan()
{
  console.log('Starting Function sendFailingStudentReportsToEthan()');
  var currentDate = new Date();
  var previousMonth = currentDate.getMonth() - 1;
  //var previousMonth = - 1 ;
  if (previousMonth == -1)
  {
    previousMonth = 11;
  }
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  emailStudentToContinueFrom = 'StudentName';
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  emailStudentFound == false;
  for (sheet in sheets)
  {
    currentStudentName = sheets[sheet].getName();

    //console.log(1);
    if (currentStudentName == "Sheet1")
    {
      continue;
    }

    //console.log(2);
    /*
    //IF ONLY TRYING TO DO ONE
    if (currentStudentName != oneStudentToExport)
    {
      continue;
    }
    */
    // CONTINUE FROM POINT
    /*
    if (currentStudentName != emailStudentToContinueFrom && emailStudentFound == false)
    {
      continue;
    }
    else if (currentStudentName == emailStudentToContinueFrom)
    {
      emailStudentFound = true;
    }
    */
    //console.log(3);

    sendThisReport = false;

    nowLastRow = sheets[sheet].getLastRow();

    sheetLocalArray = sheets[sheet].getRange('A1:H' + nowLastRow).getValues()

    for (h = 0; h < nowLastRow - 1; h++)
    {
      if ((sheetLocalArray[h][0] == overallLabel) && (typeof sheetLocalArray[h][1] == "number") && sheetLocalArray[h][1] < 50)
      {
        sendThisReport = true;
        break;
      }
    }
    //console.log(4);

    if (sendThisReport == true)
    {
      console.log('Sending Email for sheet : ' + sheet + ' out of ' + sheets.length);
      var ss = SpreadsheetApp.getActiveSpreadsheet()
      var ssID = ss.getId();
      //console.log("ssid : " + ssID);
      var sheetgId = sheets[sheet].getSheetId();
      //console.log("sheetgId : " + sheetgId);
      var sheetName = sheets[sheet].getName();
      //console.log("sheetName : " + sheetName);
      var token = ScriptApp.getOAuthToken();
      //var email = 'yourITemail@email.com';
      //var email = 
      var subject = currentStudentName + " Failled " + months[previousMonth];
      var body = "Hello, High Guardian of the Grades. This is a pdf score report for " + currentStudentName + ". You are recieving this email, because this pupil has been found wanting.";
      
      var url = "https://docs.google.com/spreadsheets/d/" +
        ssID +
        "/export?" +
        "format=xlsx" +
        "&gid=" + sheetgId +
        "&size=8.5x60.0" +
        "&portrait=true" +
        '&top_margin=0.2' +
        '&bottom_margin=0.2' +
        '&left_margin=0.2' +
        '&right_margin=0.2' +
        "&exportFormat=pdf";
      //console.log("url : " + url);

      var result = UrlFetchApp.fetch(url,
      {
        headers:
        {
          'Authorization': 'Bearer ' + token
        }
      });

      var contents = result.getContent();

      //email = 'yourITemail@email.com';
      email = 'yourITemail@email.com';

      MailApp.sendEmail(email, subject, body,
      {
        attachments: [
        {
          fileName: sheetName + ".pdf",
          content: contents,
          mimeType: "application//pdf"
        }]
      });

      console.log(sheetName + ' Grade Report Email sent to : ' + email);

      Utilities.sleep(3000);
    }
  }
}

function testSendEmail()
{
  console.log('Starting Function testSendEmails()');
  emailStudentFound = false;
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  var oneStudentToExport = "StudentName";
  for (sheet in sheets)
  {

    currentStudentName = sheets[sheet].getName();
    emailStudentToContinueFrom = 'StudentName';
    //console.log(1);
    if (currentStudentName == "Sheet1")
    {
      continue;
    }
    //console.log(2);
    /*
    //IF ONLY TRYING TO DO ONE
    
    if (currentStudentName != oneStudentToExport)
    {
      continue;
    }
    */
    // CONTINUE FROM POINT
    /*
    if (currentStudentName != emailStudentToContinueFrom && emailStudentFound == false)
    {
      continue;
    }
    else if (currentStudentName == emailStudentToContinueFrom)
    {
      emailStudentFound = true;
    }
    */
    //console.log(3);
    console.log('Sending Email for sheet : ' + sheet + ' out of ' + sheets.length);

    var ss = SpreadsheetApp.getActiveSpreadsheet()
    var ssID = ss.getId();
    //console.log("ssid : " + ssID);
    var sheetgId = sheets[sheet].getSheetId();
    //console.log("sheetgId : " + sheetgId);
    var sheetName = sheets[sheet].getName();
    //console.log("sheetName : " + sheetName);
    var token = ScriptApp.getOAuthToken();
    //var email = 'yourITemail@email.com';
    //var email = 
    var subject = currentStudentName + "'s Score Report for October";
    var body = "Hello Guardian of " + currentStudentName + ". Your child's score report is attached to this email as a PDF.";
    

    // Get dynamic range
    var lastRow = sheets[sheet].getLastRow();
    
    // Calculate dynamic height in inches (A4 content area + data height)
    var topMargin = 0.2;
    var bottomMargin = 0.2;
    var a4Width = 8.27;  // A4 width in inches (fixed)
    //var a4Height = 11.69;  // Standard A4 height in inches

    //var pageHeight = topMargin + totalInches + bottomMargin + 0.1;  // +0.1 buffer for safety
    var pageHeight = 0;  // +0.1 buffer for safety

    for (var row = 1; row <= lastRow; row++)
    {
      //console.log(row)
      pageHeight += sheets[sheet].getRowHeight(row);
      //console.log(pageHeight)
    }
    console.log(row)
    console.log(pageHeight)
    pageHeight /= 120;
    pageHeight += topMargin + bottomMargin + 0.1;
    console.log(pageHeight)
    

    var url = "https://docs.google.com/spreadsheets/d/" +
      ssID +
      "/export?" +
      "format=xlsx" +
      "&gid=" + sheetgId +
      "&size=8.5x" + pageHeight.toFixed(1) +
      //"&size=8.5x60.0" +
      "&portrait=true" +
      '&top_margin=0.2' +
      '&bottom_margin=0.2' +
      '&left_margin=0.2' +
      '&right_margin=0.2' +
      "&exportFormat=pdf";

    //console.log("url : " + url);
    var result = UrlFetchApp.fetch(url,
    {
      headers:
      {
        'Authorization': 'Bearer ' + token
      }
    });
    var contents = result.getContent();

    email = 'yourITemail@email.com';

    MailApp.sendEmail(email, subject, body,
    {
      attachments: [
      {
        fileName: sheetName + ".pdf",
        content: contents,
        mimeType: "application//pdf"
      }]
    });
    console.log(sheetName + ' Grade Report Email sent to : ' + email);
    Utilities.sleep(3000);
  }
}

function GetAndPrintStudentOverallGradesFromReportSheets()
{
  console.log("Starting Function GetAndPrintStudentOverallGradesFromReportSheets()")
  var currentPrintRow = 0;
  var stopNow = false;
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  nowLastRowFirstSheet = firstSheet.getLastRow();
  overallGradesLocalArray = firstSheet.getRange('V3:AA' + nowLastRowFirstSheet).getValues();
  for (sheet in sheets)
  {
    currentStudentName = sheets[sheet].getName();

    console.log(currentStudentName);
    console.log('Scanning sheet : ' + sheet + ' out of ' + sheets.length);

    if (currentStudentName == "Sheet1")
    {
      continue;
    }
    /*
    if (sheet < 60)
    {
      continue;
    }
    */
    nowLastRow = sheets[sheet].getLastRow();

    sheetLocalArray = sheets[sheet].getRange('A1:B' + nowLastRow).getValues()

    for (h = 0; h < nowLastRow - 1; h++)
    {
      //console.log("checking row : " + h);
      /*
      if(sheetLocalArray[h][0] == '0')
      {
        console.log("undefined 0 on row : " + h);
      }
      */
      if (sheetLocalArray[h][0] == courseLabel)
      {
        //console.log(1);
        //console.log("course name on row : " + h);
        //console.log(2);
        //console.log(currentStudentName);
        //console.log(3);
        //console.log(overallGradesLocalArray[currentPrintRow][0]);
        //console.log(4);
        //console.log(sheetLocalArray[h][1]);
        //console.log(5);
        //console.log(overallGradesLocalArray[currentPrintRow][4]);
        //console.log(6);
        overallGradesLocalArray[currentPrintRow][0] = currentStudentName;
        //console.log(7);
        overallGradesLocalArray[currentPrintRow][4] = sheetLocalArray[h][1];
        //console.log(8);
        //console.log(currentStudentName);
        //console.log(9);
        //console.log(overallGradesLocalArray[currentPrintRow][0]);
        //console.log(10);
        //console.log(sheetLocalArray[h][1]);
        //console.log(11);
        //console.log(overallGradesLocalArray[currentPrintRow][4]);
        //console.log(12);
        for (j = h; j < nowLastRow - 1; j++)
        {
          //console.log("checking row : " + j);
          /*
          if(sheetLocalArray[j][0] == '0')
          {
            console.log("undefined 0 on row: " + j);
          }
          */
          if (sheetLocalArray[j][0] == overallLabel)
          {
            //console.log("overall grade on row : " + j);
            overallGradesLocalArray[currentPrintRow][5] = sheetLocalArray[j][1];
            currentPrintRow += 1;
            break;
          }
        }
      }
      /*
      if (currentStudentName == "Irene Chen" && h > 310)
      {
        stopNow = true;
        break;

      }
      */
    }
    // IF WANT LIMIT
    /*
    if ( sheet == 3)
    {
      break;
    }
    */
    if ( stopNow == true)
    {
      break;
    }
  }
  //overallGradesLocalArray = firstSheet.getRange('V3:H' + nowLastRow).getValues();
  firstSheet.getRange('V3:AA' + nowLastRowFirstSheet).setValues(overallGradesLocalArray);
}

function SetTextPositionToLeftMostColumnIfNot()
{
  console.log('Starting Function SetTextPositionToLeftMostColumnIfNot()');
  if (alpha.indexOf(textPositions[currentStudentName][0]) > 1)
  {
    textPositions[currentStudentName][0] = 'A';
    Down(6);
    console.log("Text Position reset left to first column")
  }
}

function FindLongestSheet()
{
  var currentNumberOfRows = 0;
  var highestNumberOfRows = 0;
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() == 'Sheet1')
    {
      continue;
    }

    currentNumberOfRows = sheets[sheet].getLastRow()

    if (currentNumberOfRows > highestNumberOfRows)
    {
      highestNumberOfRows = currentNumberOfRows;
    }
  }
  console.log("Highest Number of Rows in a sheet is : " + highestNumberOfRows)
}

function clearFirstSheetTextPositions()
{
  console.log('Starting Function clearFirstSheetTextPositions()');
  firstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  firstSheet.getRange('A1:C202').clearContent();
}

function deleteEmptySheets()
{
  console.log('Starting Function deleteEmptySheets()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() == 'Sheet1')
    {
      continue;
    }
    if (sheets[sheet].getRange("A1").isBlank())
    {
      console.log(sheets[sheet].getName() + " sheet is empty, so this sheet has now been deleted")
      var ss = SpreadsheetApp.getActive();
      ss.deleteSheet(sheets[sheet]);
    }
  }
}

function deleteAllStudentSheets()
{
  console.log('Starting Function deleteAllStudentSheets()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() == 'Sheet1')
    {
      continue;
    }
    console.log(sheets[sheet].getName() + " sheet deleted")
    var ss = SpreadsheetApp.getActive();
    ss.deleteSheet(sheets[sheet]);
  }
}

function resetTallyScores()
{
  console.log('Starting Function resetTallyScores()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    //IF SPECIFIC STUDENT
    currentSheetName = sheets[sheet].getName();
    if (currentSheetName == 'Sheet1' || (onlyOneStudent == true && currentSheetName != specificStudentName))
    {
      console.log('skip : ' + currentSheetName);
      continue;
    }

    nowLastRow = sheets[sheet].getLastRow();
    numbersList = sheets[sheet].getRange('A1:H' + nowLastRow).getValues();
    for (let u = 1; u < nowLastRow; u++)
    {
      if (numbersList[u][0] == scoreWeightLabel)
      {
        //sheets[sheet].getRange(u + 1,2).setValue('no input yet');
        numbersList[u + 1][1] = 'no input yet';
        //sheets[sheet].getRange(u + 1,3).setValue('');
        numbersList[u + 1][2] = '';
        //sheets[sheet].getRange(u + 1,4).setValue('');
        numbersList[u + 1][3] = '';
      }
      if (numbersList[u][0] == overallLabel)
      {
        //sheets[sheet].getRange(u,2).setValue('no input yet');
        numbersList[u][1] = 'no input yet';
        //sheets[sheet].getRange(u + 1,3).setValue('');
        numbersList[u][2] = '';
        //sheets[sheet].getRange(u + 1,4).setValue('');
        numbersList[u][3] = '';

      }
    }
    sheets[sheet].getRange('A1:H' + nowLastRow).setValues(numbersList);
    console.log('Reset Tally Numbers for : ' + numbersList[0][1])
  }
}

function clearTallySideScores()
{
  console.log('Starting Function clearTallySideScores()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (sheet in sheets)
  {
    //IF SPECIFIC STUDENT
    currentSheetName = sheets[sheet].getName();
    if (currentSheetName == 'Sheet1' || (onlyOneStudent == true && currentSheetName != specificStudentName))
    {
      console.log('skip : ' + currentSheetName);
      continue;
    }

    nowLastRow = sheets[sheet].getLastRow();
    numbersList = sheets[sheet].getRange(1, 1, nowLastRow, 8).getValues();
    for (let u = 1; u < nowLastRow; u++)
    {
      if (numbersList[u][0] == scoreWeightLabel)
      {
        //sheets[sheet].getRange(u + 1,2).setValue('no input yet');
        //numbersList[u + 1][1] = 'no input yet';
        //sheets[sheet].getRange(u + 1,3).setValue('');
        numbersList[u + 1][2] = '';
        //sheets[sheet].getRange(u + 1,4).setValue('');
        numbersList[u + 1][3] = '';
      }
      if (numbersList[u][0] == overallLabel)
      {
        //sheets[sheet].getRange(u,2).setValue('no input yet');
        //numbersList[u][1] = 'no input yet';
        //sheets[sheet].getRange(u + 1,3).setValue('');
        numbersList[u][2] = '';
        //sheets[sheet].getRange(u + 1,4).setValue('');
        numbersList[u][3] = '';

      }
    }
    console.log("1, 1, " + nowLastRow + ", 8");
    //console.log(nowLastRow - 1);
    //console.log(nowLastRow);
    console.log(numbersList.length);
    console.log('student name : ' + numbersList[0][1]);
    //sheets[sheet].getRange('A1:H' + nowLastRow).setValues(numbersList);
    sheets[sheet].getRange(1, 1, nowLastRow, 8).setValues(numbersList);
    console.log('cleared side Tally Numbers for : ' + currentSheetName);
  }
}

//This function can be run on its own if you want to clear all of the sheets of all their content
function clearAllSheets()
{
  console.log('Starting Function clearAllSheets()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheet != 0)
    {
      sheets[sheet].clearContents();
      //sheets[sheet].clearFormats();
      sheets[sheet].autoResizeRows(1, 1);
    }
  }
  console.log('All sheets cleared');
}

var specificStudentFound = false;

function makeStudentList()
{
  
  console.log('Starting Function makeStudentList()');
  const response = Classroom.Courses.list();
  const courses = response.courses;
  if (!courses || courses.length === 0)
  {
    console.log('No courses found.');
    LabelandData('No Courses Found', '')
    return;
  }
  if (onlyOneStudent == true)
  {
    for (const course of courses)
    {
      if (specificStudentFound == true)
      {
        break;
      }
      //ENABLE TO ONLY USE ACTIVE CLASSES
      if (course.courseState != 'ACTIVE')
      {
        continue;
      }
      const studentresponse = Classroom.Courses.Students.list(course.id)
      const students = studentresponse.students;

      if (!students) continue;

      for (student in students)
      {
        if (students[student].profile.name.fullName == specificStudentName)
        {
          eachstudent = students[student].profile.name.fullName
          //console.log('each student : ' + eachstudent)
          //console.log('type of each student : ' + typeof(eachstudent))
          allstudents[eachstudent] = students[student].profile.id;
          console.log('Found specific student for one-student grade retrieval. Added student name and id to list of this one student.');
          specificStudentFound = true;
          break;
        }
      }
          
    }
  }
  else
  {
    // Print the course names and IDs of the courses
    for (const course of courses)
    {
      //ENABLE TO ONLY USE ACTIVE CLASSES
      if (course.courseState != 'ACTIVE')
      {
        continue;
      }
      const studentresponse = Classroom.Courses.Students.list(course.id)
      const students = studentresponse.students;

      if (!students) continue;
      
      for (student in students)
      {
        if (!Object.values(allstudents).includes(students[student].profile.id))
        {
          eachstudent = students[student].profile.name.fullName
          //console.log('each student : ' + eachstudent)
          //console.log('type of each student : ' + typeof(eachstudent))
          allstudents[eachstudent] = students[student].profile.id;
        }
      }
          
    }
  }
  
  console.log('Make Students List Done');
}

function sortAllStudents()
{
  console.log('Starting Function sortAllStudents()');
  const sortedallstudents = Object.keys(allstudents).sort().reduce(
    (obj, key) =>
    {
      obj[key] = allstudents[key];
      return obj;
    },
    {}
  );
  /*
  for (child in sortedallstudents) {
    //console.log(child);
  }
  console.log('allstudents');
  console.log(sortedallstudents);
  */
  allstudents = sortedallstudents;
  console.log('Sort All Students Done');
  //console.log("allstudents['StudentName'] : " + allstudents["StudentName"]);
}

function makeSheetTabForEachStudent()
{
  console.log('Starting Function makeSheetTabForEachStudent()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  if (onlyOneStudent == true)
  {
    for (sheet in sheets)
    {
      if (sheets[sheet].getName() != 'Sheet1')
      {
        sheetNames.push(sheets[sheet].getName());
      }
      if (sheets[sheet].getName() == specificStudentName)
      {
        sheets[sheet].clear();
      }
    }
  }
  else
  {
    for (sheet in sheets)
    {
      if (sheets[sheet].getName() != 'Sheet1')
      {
        sheetNames.push(sheets[sheet].getName());
        sheets[sheet].clear();
      }
    }
  }
  console.log('All sheets cleared');
  for (kid in allstudents)
  {
    if (!sheetNames.includes(kid))
    {
      //make sheet and write name and id
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(kid);
      currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(kid);
      currentStudentSheet.getRange("A1").setValue("Student Name");
      currentStudentSheet.getRange("B1").setValue(kid);
      currentStudentSheet.getRange("A2").setValue("Student ID");
      currentStudentSheet.getRange("B2").setValue(allstudents[kid]);
      console.log(kid);
    }
    else
    {
      //just insert name and id to existing student sheet
      currentStudentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(kid);
      currentStudentSheet.getRange("A1").setValue("Student Name");
      currentStudentSheet.getRange("B1").setValue(kid);
      currentStudentSheet.getRange("B1").setFontWeight("bold");
      currentStudentSheet.getRange("A2").setValue("Student ID");
      currentStudentSheet.getRange("B2").setValue(allstudents[kid]);
    }
    textPositions[kid] = ['A', 4];

  }
  //console.log(textPositions);
  console.log('Make Sheet Tab For each Student Done');
}

function continueTextPositions()
{
  console.log('Starting Function continueTextPositions()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() != 'Sheet1')
    {
      sheetNames.push(sheets[sheet].getName());
    }
  }
  console.log('sheetname list created');

  for (kid in allstudents)
  {
    textPositions[kid] = ['A', 4];
  }
  //console.log(textPositions);
  console.log('Set Default text positions A and 4');
}

function setTextPositionsAsLastRow()
{
  console.log('Starting Function setTextPositionsAsLastRow()');
  sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (sheet in sheets)
  {
    if (sheets[sheet].getName() != 'Sheet1')
    {
      nowLastRow = sheets[sheet].getLastRow();
      currentSheetName = sheets[sheet].getName()
      textPositions[currentSheetName] = ['A', nowLastRow + 2];
      //console.log('textPositions[ ' + currentSheetName + ' ] : ' + textPositions[currentSheetName])
    }
  }
  console.log('Set starting text positions to: last row of each sheet + 2');
}

var currentPointsRow = 0;
var currentPointsColumn = 0;
var currentcategory;
var currentMaxPoints;
var currentAssignedPoints;
var lastNum;
var reachedEnd = false;

var nowLastColumn;
var allassignedpoints;
var allmaxpoints;
var thisweight;
var thispoints;
var overallcounter = 0;
var noCategory = false;
var pickupStudentArrived = false;
var currentSheetName;
var numberOfAssignmentsForCurrentCourse;
var overallMax = 100;
var finalScore = 100;
var scoreChanged = false;
var totalNumberOfSheets = 10;
var numbersList;

var categorieslist;

function ShowGradeCategories()
{
  const response = Classroom.Courses.list();
  const courses = response.courses;
  for (course in courses)
  {
    categorieslist = courses[course].gradebookSettings.gradeCategories
    for (cat in categorieslist)
    {
      console.log(categorieslist[cat].name);
    }
    //console.log(courses[course].gradebookSettings.gradeCategories)
    //console.log(courses[course].gradebookSettings.gradeCategories.weight)
  }
}
//console.log(allstudents);

function LabelandData(label, data)
{
  SpreadsheetApp.getActiveSheet().getRange(cell).setValue(label);
  SpreadsheetApp.getActiveSheet().getRange(cell2).setValue(data);
  IncrementDown()
}

function IncrementDown()
{
  number += 1
  cell = alpha[alphanumber] + number.toString()
  cell2 = alpha[alphanumber + 1] + number.toString()
}

function IncrementRightTop()
{
  alphanumber += 3
  number = 2
  cell = alpha[alphanumber] + number.toString()
  cell2 = alpha[alphanumber + 1] + number.toString()
}

function Up(amount)
{
  textPositions[currentStudentName][1] -= amount;
  //console.log("position = " + textPositions[currentStudentName].join(''))
}

function Down(amount)
{
  textPositions[currentStudentName][1] += amount;
  //console.log("position = " + textPositions[currentStudentName].join(''))
}

function Left(amount)
{
  textPositions[currentStudentName][0] = alpha[alpha.indexOf(textPositions[currentStudentName][0]) - amount];
  //console.log("position = " + textPositions[currentStudentName].join(''))
}

function Right(amount)
{
  textPositions[currentStudentName][0] = alpha[alpha.indexOf(textPositions[currentStudentName][0]) + amount];
  //console.log("position = " + textPositions[currentStudentName].join(''))
}
/*
print = ''
 
for (let i = 0; i < alpha.length; i++)
{
  print += "\"C" + alpha[i] + "\",";
}
console.log(print)
*/
/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e)
{
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Show sidebar', 'showSidebar')
    .addItem('Show dialog', 'showDialog')
    .addToUi();
  listCourses();
}
/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e)
{
  onOpen(e);
}
/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar()
{
  var ui = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle(SIDEBAR_TITLE)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}
/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog()
{
  var ui = HtmlService.createTemplateFromFile('Dialog')
    .evaluate()
    .setWidth(400)
    .setHeight(190)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}
/**
 * Returns the value in the active cell.
 *
 * @return {String} The value of the active cell.
 */
function getActiveValue()
{
  // Retrieve and return the information requested by the sidebar.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return cell.getValue();
}
/**
 * Replaces the active cell value with the given value.
 *
 * @param {Number} value A reference number to replace with.
 */
function setActiveValue(value)
{
  // Use data collected from sidebar to manipulate the sheet.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  cell.setValue(value);
}

// If you image is stored in Google drive
// You can retrieve image as blob with DriveApp
function getImageUrlFromDriveApi()
{
  const image = Drive.Files.get('1VkTPKsgu1nF5CUQXkct5izQ4E1q9epSI');

  // From image JSON object we get a thumbnail link
  // Thumbnail link has default size set at the end '=s + size'
  // If we want to change size, split the end of url and add desired size
  // const thumbnail = image.thumbnailLink.split("=")[0] + "=s200";
  const thumbnail = image.thumbnailLink;
  return thumbnail;
}

// Image builder method
// Inserts image inside a cell as a value
function insertImageToCellWithImageBuilder(sheet, url, row, col)
{
  let image = SpreadsheetApp
    .newCellImage()
    .setSourceUrl(url)
    .build();
  sheet.getRange(row, col).setValue(image);
}

/**
 * Executes the specified action (create a new sheet, copy the active sheet, or
 * clear the current sheet).
 *
 * @param {String} action An identifier for the action to take.
 */
function modifySheets(action)
{
  // Use data collected from dialog to manipulate the spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  if (action == "create")
  {
    ss.insertSheet();
  }
  else if (action == "copy")
  {
    currentSheet.copyTo(ss);
  }
  else if (action == "clear")
  {
    currentSheet.clear();
  }
}