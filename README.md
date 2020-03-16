# Assessment-Data

## Table of Contents
  * [Introduction](#introduction)
  * [Technologies](#technologies)
  * [Setup](#setup)
  * [Features](#features)
  
## Introduction
Quick, automated grading, and grade analysis for targeted interventions in a teacher and student-friendly medium. For collaborative teams (CT) with common assessments (CA) who want to compare student performance for actionable data dialogues.

## Technologies
#### Project created with:
  * Apps Scripts (V8 runtime, 03-2020)

#### Project requires:
  * Google Sheets
  
## Setup
#### To run this project:
  * Create a spreadsheet with two sheets
  * Copy the script into its script editor (tools > script editor)

#### The code assumes:
  * Two sheets called "Outline" and "Students" (in that order)
  * "Students" sheet is populated with student period, ID, first, and last name starting on row 4.
    * Recommend pulling this from another sheet so the info auto-populates as you get new students
  * The "Outline" Sheet has
    * A drop-down list of numbers of formative assessments (ex: numbers 1-4) at column B
    * A formatted table for mapping out each assessment with formulas starting at column C

## Features
  * Autocalculate common assessment total
  * Autofill and format assessment sheets
  * Autocalculate student grade
  * Sort student performance by teacher
  * View performance on each section
#### To do
  * *Sort student performance by teacher*
  * *Data analysis (who got best results)*
  * *Customizable colour (to distinguish units?)*
  * *Compare formative and summative*
