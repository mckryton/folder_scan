@d-folder_scan
Feature: update list

    As a user I want to let the macro add only new file names to my list, so
    that the list does not contain duplicates. In addition I want save the
    filenames as links, so that I can open the files by clicking the links in
    my list.

  Scenario: list is empty
    Given an empty list
    When I run the macro
    Then file names will be added just below the headline of my list

  Scenario: list is not empty
    Given the list contains at least one filename
    When I run the macro
    Then file names will be added just below the headline of my list

  Scenario: file is new
    Given the folder contains a file
     And the file is not in the list
    When I run the macro
    Then the name of the file will be appended at the end of the list
     And a hyperlink from the list entry will refer to the location of the filename
     And the current date will be set beside the filenames

  Scenario: file exists
    Given the folder contains a file
     And the file is not in the list
    When I run the macro
    Then the name of this file will be ignored

  Scenario: copy list
    Given a list refers to a folder
    When I copy the list to a new sheet into the same Excel file
     And I change the folder path to another folder in the new sheet
    Then the macro will the use this another folder to update the new sheet
     And the macro will use the original folder when updating the original sheet
