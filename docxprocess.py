import aspose.words as aw
import re

def grab_image_and_text(input_docx_file, image_sav_loc, img_prefix):
    # Load the Word document
    document = aw.Document(input_docx_file)
    pattern = r'[^a-zA-Z0-9 @*._$-]'

    pattern2 = r'Evaluation Only. Created with Aspose.Words. Copyright 2003-2023 Aspose Pty Ltd.'

    # Initialize an array to store the extracted text between images
    text_between_images = []
    img_file_list = []
    # Intialize counter and text
    count = 0
    text = ""
    # Iterate through the nodes in the document
    for node in document.get_child_nodes(aw.NodeType.ANY, True):
        if node.node_type == aw.NodeType.SHAPE:
            # We found a shape, so set the inside_text flag to False
            shape = node.as_shape()
            if shape.has_image:
                # set image file's name
                imageFileName = f"{img_prefix}{count}{aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)}"
                img_file_list.append(imageFileName)
                # save image
                shape.image_data.save(image_sav_loc + imageFileName)
                count = count + 1
                text_between_images.append(text)
                text = ""            
        elif node.node_type == aw.NodeType.RUN:
            # We found text in a Run node
            run = node.as_run()
            run_text = run.get_text()
            # print(run_text)
            text = str(text + " " +str(run_text))
            

    # Get the documents after the last image also
    text_between_images.append(text)
    
    # Remove the specified number of starting words
    search_string = "Aspose.Words"
    if len(text_between_images[0]) > 80 and search_string in text_between_images[0]:
        print("Found the mention")
        substring = text_between_images[0][80:]
        text_between_images[0] = substring

    # Remove last characters from Aspose
    for idx in range(len(text_between_images)):
        # cleaned_string = re.sub(pattern, '', text_between_images[idx])
        # cleaned_string2 = re.sub(pattern2, '', cleaned_string)
        text_between_images[idx] = text_between_images[idx]
        if search_string in text_between_images[idx]:
            # print("Found at: " + text_between_images[idx])    
            if len(text_between_images[idx]) > 142:
                substring = text_between_images[idx][:-142]
                text_between_images[idx] = substring
                # print("After removal: " + text_between_images[idx])
     
    # Remove the last element from the list
    img_file_list.pop(len(img_file_list)-1)

    # # Printing for testing    
    # for idx in range(len(text_between_images)):
    #     print(text_between_images[idx] + "\n")

    return text_between_images, img_file_list

