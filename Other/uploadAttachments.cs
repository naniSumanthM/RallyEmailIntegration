// System Libraries
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Web;

// Rally REST API Libraries
using Rally.RestApi;
using Rally.RestApi.Response;

namespace RestExample_CreateAttachment
{
    class Program
    {
        static void Main(string[] args)
        {

			string storyReference = https://rally1.rallydev.com/slm/webservice/v2.0/hierarchicalrequirement/70836533324";
		
            //Process - Attaching the image to the user story
			
            // Read In Image Content
            String imageFilePath = "C:\\Users\\username\\";
            String imageFileName = "image1.png";
            String fullImageFile = imageFilePath + imageFileName;
            Image myImage = Image.FromFile(fullImageFile);

            // Convert Image to Base64 format
            string imageBase64String = ImageToBase64(myImage, System.Drawing.Imaging.ImageFormat.Png);

            // Length calculated from Base64String converted back
            var imagenumberbytes = convert.frombase64string(imagebase64string).length;

            // DynamicJSONObject for AttachmentContent
            DynamicJsonObject myAttachmentContent = new DynamicJsonObject();
            myAttachmentContent["Content"] = imageBase64String; //string 

            try
            {
                CreateResult myAttachmentContentCreateResult = restApi.Create("AttachmentContent", myAttachmentContent); //dynamicJsonObj, myAttachmentContent
                //Extra code
				String myAttachmentContentRef = myAttachmentContentCreateResult.Reference;
                Console.WriteLine("Created: " + myAttachmentContentRef);

				
                // DynamicJSONObject for Attachment Container
                DynamicJsonObject myAttachment = new DynamicJsonObject();
                myAttachment["Artifact"] = storyReference;
                myAttachment["Content"] = myAttachmentContentRef;
                myAttachment["Name"] = "AttachmentFromREST.png";
                myAttachment["Description"] = "Attachment Desc";
                myAttachment["ContentType"] = "image/png";
                myAttachment["Size"] = imageNumberBytes;
              //myAttachment["User"] = myUserRef;

                CreateResult myAttachmentCreateResult = restApi.Create("Attachment", myAttachment);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unhandled exception occurred: " + e.StackTrace);
                Console.WriteLine(e.Message);
            }
        }

		
		//ImageToBase64()
        // Converts image to Base 64 Encoded string
        public static string ImageToBase64(Image image, System.Drawing.Imaging.ImageFormat format)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                image.Save(ms, format);
                // Convert Image to byte[]
                byte[] imageBytes = ms.ToArray();
				
                // Convert byte[] to Base64 String
                string base64String = Convert.ToBase64String(imageBytes);
                
				return base64String;
            }
        }
    }
}



