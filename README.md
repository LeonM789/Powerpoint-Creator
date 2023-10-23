<h1 align="center">
  <br>
  <a href="https://leonmarx.de"><img src="https://leonmarx.de/other_pictures/github_logo.jpeg" alt="Markdownify" width="200"></a>
  <br>
  PowerPoint Creator
  <br>
</h1>

<h4 align="center"><b><u>Automated PowerPoint creator for the <a href="https://hm.edu" target="_blank">HM</b></u></a>.</h4>

<!---
<p align="center">
  <a href="#introduction">Introduction</a> •
  <a href="#Requirements">Requirements</a> •
  <a href="#Usage">Usage</a> •
  <a href="#Example-function">Example Function</a> •
  <a href="#related">Related</a> •
  <a href="#license">License</a>
</p>
--->
 
<br>

## 📝 Introduction
For the graduation ceremony of the Munich University of Applied Sciences, a PowerPoint presentation has to be created every year, in which each graduate gets a slide. In order not to have to laboriously create this presentation by hand, this repository contains Python code that creates it almost completely automatically.

<br>

*(For the chosen one, who now has the honor of creating this presentation for the graduation ceremony: you will find the more extensive powerpoint template in the FS cloud)*

<br>

## 🛠️ Requirements
This repository contains a Jupyter notebook and the notebook combined into a Python file, which require the following packages:
<ul>
  <li>os (build in)</li>
  <li>requests</li>
  <li>urllib3</li>
  <li>bs4</li>
  <li>pptx</li>
</ul>
You can easily install them via pip by running the first cell in the notebook or by using the terminal.

<br>
<br>

**It could be the case that you need an older version of Python for the download section in the notebook (e.g. 3.9.7) and a newer one (e.g. 3.11.1) for the creation part (at least I had some problems with the python versions).**

<br>

## 💡 Usage
Note that this code is designed exactly for the template provided and changes must always be made in the PowerPoint slide master *AND* in the code to avoid errors.

<br>

1. Make sure you have a link and access to a file server where all the files (txt and jpg) are linked
2. Download the template and the notebook or the Python file
3. Install all required packages by running the first cell or by using the terminal
4. Give the variable url the link to the file register
5. Adjust the code or your data and the powerpoint template to your needs <br>
    5.5. Run the Python file by using the terminal (cd *your/folder/path* , pyhton3 pp_creator.py) <br>
    OR <br>
    5.5. Run the "download cell" in the notebook to download all the data to the folder where the notebook is located <br>
    5.6Run the "create cell" to create the powerpoint <br>
6. Adjust the images in the powerpoint that are not well fitted <br>
7. Add your own slides that you need

<br>

## 🗂️ Example Data
The txt files for which my code is written are build like this: 

```
Nachname: *first name*     Vorname: *last name*    Highlight: *bla bla bla* Hobbies: *bla bla bla*     Studiengang: *major* 
```
... and are named like this (!important):
```
GatesBill.txt
GatesBill.jpg
```

The code and the template covers the following situations:
1. There is only a txt for the student and the only information in it are the first name and the last name
2. There is also a picture of the student
3. In addition to 2., the student provides information about his/her highlights and/or hobbies 
3. In addition to his name, the student only provides his highlights/hobbies (no picture)

<br>
<!---
## 🧡 You may also like...

- [FlappyBird AI](https://github.com/LeonM789/FlappyBirdAI.git) - A neural network in Python

<br>
--->

## ⚖️ License

This project is licensed under the MIT License - see the `LICENSE` file for details.

<br>

---

> [leonmarx.de](https://www.leonmarx.de) &nbsp;&middot;&nbsp;
> GitHub [@LeonM789](https://github.com/LeonM789) 
