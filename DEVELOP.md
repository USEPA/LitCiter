# Instructions for development

To create a development environment:
* Clone this repo, navigate to the folder
* `conda create -n litciter nodejs=16`
* `conda activate litciter`
* `npm install`
* `npx office-addin-usage-data off`
* `npm start`

To deolpy:
* `npm run build`
* Distribute `manifest.xml`, changing urls
* Host files in `dist`

See [here](https://github.com/OfficeDev/generator-office) for more details.