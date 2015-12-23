/*--------------------------------------------------------------------------
Project		: Word Automation
Component	: Automation
Module		: Word Automation
Description	: This class is used for word automation using dynamic. 
              Aupport almost all the versions. 
Log
Date				Author							Comment
08-Nov-2011			Prathapachandran				Created
--------------------------------------------------------------------------*/
using System;
using System.IO;
using System.Reflection;

namespace Automation
{
	/// <summary>
	/// Cass for automating word.
	/// </summary>
	public class WordAutomation : IDisposable
	{
        #region Member Variables
        //Dynamic object for word application
        private dynamic _wordApplication;
        //Dynamic object for word document
        private dynamic _wordDoc;
        #endregion 

        #region Member Functions

	    /// <summary>
		/// Function to create a word document.
		/// </summary>
		/// <param name="fileName">Name of word</param>
		/// <param name="isReadonly">open mode</param>
		/// <returns>If word document exist, return word doc object
		/// else return null</returns>		
		public void CreateWordDoc(object fileName, bool isReadonly)
		{
	        if (!File.Exists(fileName.ToString()) || _wordApplication == null) return;
	        //object readOnly = isReadonly;
	        object isVisible = true;
	        object missing = Missing.Value;
	        // Open a given Word document.
	        _wordDoc = _wordApplication.Documents.Open(fileName, missing,
	            isReadonly, missing, missing, missing,
	            missing, missing, missing, missing,
	            missing, isVisible);
		}
		/// <summary>
		/// Function to create a word application
		/// </summary>
		/// <returns>Returns a word application</returns>		
		public void CreateWordApplication()
		{
            const string message = "Failed to create word application. Check whether"
                                   + " word installation is correct.";	
			var wordType = Type.GetTypeFromProgID("Word.Application");	            
			if (wordType == null)
			{
                throw new Exception(message);
			}
            _wordApplication = Activator.CreateInstance(wordType);
            if (_wordApplication == null)
			{
                throw new Exception(message);
			}
		}
		/// <summary>
		/// Function to close a word document.
		/// </summary>
		/// <param name="canSaveChange">Need to save changes or not</param>
		/// <returns>True if successfully closed.</returns>		
		public bool CloseWordDoc(bool canSaveChange)
		{
		    if (_wordDoc == null) return false;
		    object saveChanges;				
		    if (canSaveChange)
		    {
		        saveChanges = -1; // Save Changes
		    }
		    else
		    {
		        saveChanges = 0; // No changes
		    }
		    _wordDoc.Close(saveChanges);
		    //InvokeMember("Close",wordDocument,new object[]{saveChanges});	
		    _wordDoc = null;
		    return true;
		}
		/// <summary>
		/// Function to close word application
		/// </summary>
		/// <returns>True if successfully closed</returns>		
		public bool CloseWordApp()
		{
			bool isSuccess = false;
			if (_wordApplication != null)
			{								
				object saveChanges = -1; // Save changes
                _wordApplication.Quit(saveChanges);
                _wordApplication = null;
				isSuccess = true;
			}
			return isSuccess;
		}
		/// <summary>
		/// Function to get the word count from the document. Entire word should 
		/// match.
		/// </summary>
		/// <param name="word">Word to search</param>
		/// <returns>Count of the word</returns>		
		public int GetWordCount(string word)
		{
           // object wordDoc = _wordDoc;
			int count = 0;
			do
			{
				if (_wordDoc == null)
				{
					break;
				}
				if (word.Trim().Length == 0)
				{
					break;
				}
                _wordDoc.Activate();
             //   dynamic content = _wordDoc.Content;

				// Get the count from direct text inside the document.
                count += GetCountFromRange(_wordDoc.Content, word);	
                int rangeCount = _wordDoc.Comments.Count;

			    const int x = 1;
			    for(; x <= rangeCount;)
				{
                    count += GetCountFromRange(_wordDoc.Comments.Item(x), word);
					break;
				}
// duplicate code
                rangeCount = _wordDoc.Sections.Last.Headers.Count;
				for (int i = 1; i <= rangeCount; i++)
				{                  
                    count += GetCountFromRange(
                        _wordDoc.Sections.Last.Headers.Item(i).Range, word);
				}
                rangeCount = _wordDoc.Sections.Last.Footers.Count;

				for (int i = 1; i <= rangeCount; i++)
				{
                    count += GetCountFromRange(
                        _wordDoc.Sections.Last.Footers.Item(i).Range, word);
				}
                rangeCount = _wordDoc.Shapes.Count;

				for (int i = 1; i <= rangeCount; i++)
				{                   
                    dynamic textFrame = _wordDoc.Shapes.Item(i).TextFrame;
                    int hasText = textFrame.HasText;
                    //duplicate code
					if (hasText < 0)
					{
                        count += GetCountFromRange(textFrame.TextRange, word);
					}
				}			
			}
			while(false);
			return count;
		}		
		/// <summary>
		/// Function to get word count from a given range.
		/// </summary>
		/// <param name="range">Range</param>
		/// <param name="word">Word to search</param>
		/// <returns>Count of words</returns>
        private int GetCountFromRange(dynamic range,string word)
        {
            int count = 0;
		    object item = 1; // Goto Page
            object whichItem = 1;// First page	
            _wordDoc.Goto(item, whichItem);            
            dynamic find = range.Find;
            find.ClearFormatting();
            find.Forward = true;
            find.Text = word;
            find.MatchWholeWord = true;
            find.Execute();
            bool found = find.Found;
            while (found)
            {
                ++count;
                find.Execute();
                found = find.Found;
            }
            return count;
        }
		/// <summary>
		/// Function to find and replace a given text.
		/// </summary>
		/// <param name="findText">Text for finding</param>
		/// <param name="replaceText">Text for replacing</param>
		/// <returns>True if successfully replaced.</returns>
		public bool FindReplace(
			string findText, string replaceText)
		{
            
			var isSuccess = false;
			do
			{
				if (_wordDoc == null)
				{
					break;
				}
				if (_wordApplication == null)
				{
					break;
				}
				if (replaceText.Trim().Length == 0)
				{
					break;
				}
				if (findText.Trim().Length == 0)
				{
					break;
				}

                ReplaceRange(_wordDoc.Content, findText, replaceText);

                int rangeCount = _wordDoc.Comments.Count;
				for(int i = 1; i <= rangeCount; i++)
				{
                    ReplaceRange(_wordDoc.Comments.Item(i).Range, 
                        findText, replaceText);					
				}
                //duplicate code
                rangeCount = _wordDoc.Sections.Last.Headers.Count;
				for (int i = 1; i <= rangeCount; i++)
				{
                    ReplaceRange(_wordDoc.Sections.Last.Headers.Item(i).Range, 
                        findText, replaceText);
				}
                rangeCount = _wordDoc.Sections.Last.Footers.Count;
				for (int i = 1; i <= rangeCount; i++)
				{
                    ReplaceRange(_wordDoc.Sections.Last.Footers.Item(i).Range, 
                        findText, replaceText);
				}

                rangeCount = _wordDoc.Shapes.Count;
				for (int i = 1; i <= rangeCount; i++)
				{
                    dynamic textFrame = _wordDoc.Shapes.Item(i).TextFrame;
					int hasText = textFrame.HasText;
	//			    duplicatecode
					if (hasText < 0)
					{
                        ReplaceRange(textFrame.TextRange, findText, replaceText);
					}
				}
				isSuccess = true;
			}
			while(false);
			return isSuccess;
		}
		/// <summary>
		/// Replace a word with in a range.
		/// </summary>
		/// <param name="range">Range to replace</param>
		/// <param name="findText">Text to find</param>
		/// <param name="replaceText">Text to replace</param>
		private void ReplaceRange(dynamic range, 
			string findText, string replaceText)
		{
			object missing = Missing.Value;
            _wordDoc.Activate();
			object item = 1;
			object whichItem = 1;
            _wordDoc.GoTo(item, whichItem);
			object replaceAll = 2;
		    dynamic find = range.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Execute(findText, false, true,
                                missing, missing, missing, true, missing, missing
                                , replaceText, replaceAll);														
		}
		#endregion
        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _wordApplication = null;
                    _wordDoc = null; //dispose managed ressources
                }
            }
            //dispose unmanaged ressources
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
	  
	}
}