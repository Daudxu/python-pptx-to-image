from spire.presentation.common import *
from spire.presentation import *

# Create a Presentation object
presentation = Presentation()

# Load a PPT or PPTX file
presentation.LoadFromFile("C:/Users/Administrator/Desktop/input.pptx")

# Loop through the slides in the presentation
for i, slide in enumerate(presentation.Slides):

    # Specify the output file name
    fileName ="Output/ToImage_ + str(i) + ".png"
    # Save each slide as a PNG image
    image = slide.SaveAsImage()
    image.Save(fileName)
    image.Dispose()

presentation.Dispose()
