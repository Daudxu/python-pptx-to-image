import aspose.slides as slides
import aspose.pydrawing as drawing

# Load presentation
pres = slides.Presentation("test1.pptx")

# Loop through slides
for index in range(pres.slides.length):
    # Get reference of slide
    slide = pres.slides[index]

    # Define scaling
    scaleX = 2
    scaleY = 2

    # Save as PNG
    slide.get_thumbnail(scaleX, scaleY).save("slide_{i}.png".format(i = index), drawing.imaging.ImageFormat.png)
