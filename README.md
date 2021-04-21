# Auto Slides
Script for generating presentation slides from a given topic.

## Features:
  - Finds the term defenitions
  - Approximates the best defenition
  - Finds appropriate pictures
  - Compiles everything to .pptx presentation

### Installation:
```sh
git clone https://github.com/Azarattum/AutoSlides.git
cd AutoSlides
pip install -r requirements.txt
```

### Usage: 
```sh
gen_slydes.py [-h] [--search SEARCH] [--author AUTHOR] [--no-best] [--no-sources] term
```

### Arguments:
| Long         | Short | Description                                         |
| ------------ | ----- | --------------------------------------------------- |
| --help       | -h    | Show help message                                   |
| --search     | -s    | Specify additional search parameters                |
| --author     | -a    | Add the presentation author                         |
| --no-best    |       | Exclude the best defenition slide from presentation |
| --no-sources |       | Exclude sources slide from presentation             |

### Third-party libraries:
* [python-pptx](https://github.com/scanny/python-pptx) - Create Open XML PowerPoint documents in Python.
* [requests](https://github.com/psf/requests) - A simple, yet elegant HTTP library.
