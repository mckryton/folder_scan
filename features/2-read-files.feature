@d-folder_scan
Feature: read files

  As a user I want to read all files from a given folders, so that I can get
  a list from them.


  Scenario: folder is empty
    Given no folder is set as default in myExcel file
    When I run the macro
    Then I will see I dialog asking me to choose a folder


  Scenario: folder is not accessible
    Given a folder is set as default in my Excel file
     And the path of this folder is not available
    When I run the macro
    Then I will see an error message that the folder is not accessible


  Scenario: folder contains files
    Given a folder is set as default in my Excle file
      And the folder is accessible
      And the folder contains files
    When I run the macro
    Then I will see all those filenames in my list
