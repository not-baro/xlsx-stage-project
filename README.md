# xlsx-stage-project
I made this tool in sept-oct-nov 2020 as a stage project in UNIMIB computer science course. The tool takes in input an xlsx file, specifically set up with columns and names by the professor, and it catalogues some important informations in order to make life easier to the professor to select most important phrases and incentivate phrases' discussion between the students during the course. It also counts how many phrases every student inserted in the google form, so that the professor can easily see how much effort every student put in for the work. The tool comes with a simple web-app where you can upload the xlsx file and receive back the output file.

flask_app.py contains the final version of the script alongside the webapp, if you want to edit/use the script without the webapp, you may just delete the def process(uploaded): function and fix the indentation of the content inside the function. Also delete everything that has to do with Flask framework.

Download the archive and launch 'pip install -r requirements' to install all the needed dependencies.
