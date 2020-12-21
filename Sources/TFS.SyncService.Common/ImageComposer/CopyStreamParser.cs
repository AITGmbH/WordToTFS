namespace AIT.TFS.SyncService.Common.ImageComposer
{
    #region Usings
    using System;
	using System.Collections.Generic;
	using System.Diagnostics.CodeAnalysis;
	using System.IO;
	using System.Text.RegularExpressions;
	using Microsoft.Office.Interop.Word;
    #endregion

    /// <summary>
    /// Parser to find every image file definied in vml and html part of html stream created by copy of word table cell content.
    /// </summary>
    public class CopyStreamParser
	{
		#region Prive fields

		private List<CopyStreamImage> imageFiles = new List<CopyStreamImage>();

		#endregion Prive fields

		#region Private properties

		/// <summary>
		/// Gets all found image files.
		/// </summary>
		/// <value>
		/// The found image files.
		/// </value>
		public IList<CopyStreamImage> ImageFiles
		{
			get
			{
				return imageFiles;
			}
		}

		#endregion Private properties

		#region Public methods

		/// <summary>
		/// Parses the specified stream and find images.
		/// </summary>
		/// <param name="stream">The stream to parse and find images.</param>
		/// <param name="range">Corresponding range in word.</param>
		/// <returns>Repaired stream.</returns>
		public string ParseAndRepairForCopy(string stream, Range range)
		{
			this.Parse(stream, range);

			if (string.IsNullOrEmpty(stream))
			{
				return stream;
			}

			var newStream = stream;

			for (var index = this.ImageFiles.Count - 1; index >= 0; index--)
			{
				var pair = this.ImageFiles[index];
				var vmlPartFile = stream.Substring(pair.VmlFileIndex, pair.VmlFileLength);
				var htmlPartFile = stream.Substring(pair.HtmlFileIndex, pair.HtmlFileLength);
				if (vmlPartFile == htmlPartFile || pair.IsOleObject)
				{
					continue;
				}

				// Version 'replace files'
				////var vmlUri = new Uri(vmlPartFile);
				////if (!File.Exists(vmlUri.LocalPath))
				////{
				////    continue;
				////}
				////var htmlUri = new Uri(htmlPartFile);
				////if ((!File.Exists(htmlUri.LocalPath)))
				////{
				////    continue;
				////}
				////File.Copy(vmlUri.LocalPath, htmlUri.LocalPath, true);

				// Version 'replace src in img tag'
				if (pair.HtmlAltStart > pair.HtmlFileIndex)
				{
					// 'alt' part exists and it is after 'src' part
					// Replace 'alt' at first and 'src' after that
					newStream = newStream.Remove(pair.HtmlAltStart, pair.HtmlAltEnd - pair.HtmlAltStart);
					newStream = newStream.Insert(pair.HtmlAltStart, pair.AltPart);
					newStream = newStream.Remove(pair.HtmlFileIndex, pair.HtmlFileLength);
					newStream = newStream.Insert(pair.HtmlFileIndex, vmlPartFile);
				}
				else
				{
					// 'alt' part don't exists, or it is before 'src' part
					// Replace 'src' at first and 'alt' after that
					newStream = newStream.Remove(pair.HtmlFileIndex, pair.HtmlFileLength);
					newStream = newStream.Insert(pair.HtmlFileIndex, vmlPartFile);
					if (pair.HtmlAltEnd != -1)
					{
						// 'alt' part exists
						newStream = newStream.Remove(pair.HtmlAltStart, pair.HtmlAltEnd - pair.HtmlAltStart);
					}
					newStream = newStream.Insert(pair.HtmlAltStart, pair.AltPart);
				}
			}
#if DEBUG
			// Temp operation - create ald and new stream in 'b:\d' folder.
			////this.CreateBeforeFile(stream);
			////this.CreateAfterFile(newStream);
#endif // DEBUG

			return newStream;
		}

		#endregion Public methods

		#region Private methods - only for debug

#if DEBUG
		/// <summary>
		/// Gets the name of the file in 'c:\d' folder to write the html stream to the disk.
		/// </summary>
		/// <param name="before"><c>true</c> - file for 'before' html strem is required. <c>false</c> - file for 'after' html stream is required.</param>
		/// <returns>Required file, that not exists on the disk.</returns>
		[SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "At this moment not used.")]
		private static string GetFileName(bool before)
		{
			for (var index = 0; ; index++)
			{
				var fileName = (before ? "Before_" : "After_") + index;
				var fullFileName = Path.Combine(@"c:\e", fileName);
				if (!File.Exists(fullFileName))
				{
					return fullFileName;
				}
			}
		}

		/// <summary>
		/// Creates the file with 'before' html stream.
		/// </summary>
		/// <param name="stream">The stream to write.</param>
		[SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "At this moment not used.")]
		private static void CreateBeforeFile(string stream)
		{
			using (var writer = File.CreateText(GetFileName(true)))
			{
				writer.Write(stream);
			}
		}

		/// <summary>
		/// Creates the file with 'after' html stream.
		/// </summary>
		/// <param name="stream">The stream to write.</param>
		[SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "At this moment not used.")]
		private static void CreateAfterFile(string stream)
		{
			using (var writer = File.CreateText(GetFileName(false)))
			{
				writer.Write(stream);
			}
		}
#endif // DEBUG

		#endregion Private methods - only for debug

		#region Private methods

		/// <summary>
		/// The method parses the specified stream and finds images.
		/// </summary>
		/// <param name="stream">The stream to parse and find images.</param>
		/// <param name="range">Corresponding range in word.</param>
		private void Parse(string stream, Range range)
		{
			// Clear list
			this.imageFiles.Clear();

			// Find 'vml' parts with correct image
			var pattern = @"<!--\[if\sgte\svml(.+?)<!\[endif]-->";
			var vmlMatches = Regex.Matches(stream, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

			// Find 'html' parts with wrong image
			pattern = @"<!\[if\s!vml(.+?)<!\[endif]>";
			var htmlMatches = Regex.Matches(stream, pattern, RegexOptions.Singleline | RegexOptions.IgnoreCase);

			// Check pairs
			if (vmlMatches.Count != htmlMatches.Count)
			{
				return;
			}

			// Check consistency 1. pass
			for (int index = 0; index < vmlMatches.Count; index++)
			{
				var vmlMatch = vmlMatches[index];
				var htmlMatch = htmlMatches[index];
				if (vmlMatch.Index + vmlMatch.Length > htmlMatch.Index)
				{
					return;
				}
			}

			// Check consistency 2. pass
			for (int index = 0; index < vmlMatches.Count - 1; index++)
			{
				var vmlMatch = vmlMatches[index + 1];
				var htmlMatch = htmlMatches[index];
				if (htmlMatch.Index + htmlMatch.Length > vmlMatch.Index)
				{
					return;
				}
			}

			// Precompile regex to find image part
			var srcExpression = new Regex(@"(<(img\s|v:imagedata\ssrc)[^>]*>)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

			// Create parts
			for (int index = 0; index < vmlMatches.Count; index++)
			{
				var vmlMatch = vmlMatches[index];
				var htmlMatch = htmlMatches[index];
				var vmlFile = srcExpression.Matches(vmlMatch.Value);
				var htmlFile = srcExpression.Matches(htmlMatch.Value);
				if (vmlFile.Count != 1 || vmlFile.Count != htmlFile.Count)
				{
					continue;
				}
				const string ConstSrc = "src=\"";
				const string ConstAlt = "alt=\"";
				const string ConstEnd = "\"";

				// extract the file path --> it's between the src=" and the next closing "
				var vmlStart = 5 + vmlFile[0].Value.IndexOf(ConstSrc, StringComparison.OrdinalIgnoreCase);
				var vmlEnd = vmlFile[0].Value.IndexOf(ConstEnd, vmlStart, StringComparison.OrdinalIgnoreCase);
				var htmlStart = 5 + htmlFile[0].Value.IndexOf(ConstSrc, StringComparison.OrdinalIgnoreCase);
				var htmlEnd = htmlFile[0].Value.IndexOf(ConstEnd, htmlStart, StringComparison.OrdinalIgnoreCase);

				// Start and End for 'alt="....."' means whole alt part. If this part was found.
				var altStart = htmlFile[0].Value.IndexOf(ConstAlt, StringComparison.OrdinalIgnoreCase);
				var altEnd = -1;
				if (altStart == -1)
				{
					// Not found - set only start and let the 'altEnd' unchanged
					// subtract before added 5
					altStart = htmlStart - 5;
				}
				else
				{
					// Found - set end
					// Find starts after 'alt="'
					altEnd = 1 + htmlFile[0].Value.IndexOf(ConstEnd, altStart + 5, StringComparison.OrdinalIgnoreCase);
				}

				// Position = 'Position of block in stream'
				//          + 'Position of node in block'
				//          + 'Position of file in node'
				var vmlFileIndex = vmlMatch.Index + vmlFile[0].Index + vmlStart;

				// Length is simple
				var vmlFileLength = vmlEnd - vmlStart;

				// Position = 'Position of block in stream'
				//          + 'Position of node in block'
				//          + 'Position of file in node'
				var htmlFileIndex = htmlMatch.Index + htmlFile[0].Index + htmlStart;

				// Length is simple
				var htmlFileLength = htmlEnd - htmlStart;

				var isOleObject = DetermineIfImageIsOleObject(stream, vmlFileIndex);

				// Create instance
				var copyStreamImage = new CopyStreamImage(vmlFileIndex, vmlFileLength, htmlFileIndex, htmlFileLength, isOleObject);

				// Both position = 'Position of block in stream'
				//                  + 'Position of node in block'
				//                  + 'Position'
				copyStreamImage.HtmlAltStart = htmlMatch.Index + htmlFile[0].Index + altStart;
				copyStreamImage.HtmlAltEnd = altEnd == -1 ? altEnd : htmlMatch.Index + htmlFile[0].Index + altEnd;

				// Add element to the list
				imageFiles.Add(copyStreamImage);
			}

			// Don't access the shapes with '[]' operator
			int shapeIndex = 0;
			foreach (InlineShape shape in range.InlineShapes)
			{
				if (shapeIndex < this.imageFiles.Count)
				{
					this.imageFiles[shapeIndex].SetImageInformation(shape);
				}
				shapeIndex++;
			}
		}

		/// <summary>
		/// Determines if the surrounding paragraph of the begin position contains a OLEObject tag.
		/// </summary>
		/// <param name="stream">The html clipboard copy stream to search.</param>
		/// <param name="currentBeginPosition">The position from where the surrounding paragraph tag will be used.</param>
		/// <returns></returns>
		private bool DetermineIfImageIsOleObject(string stream, int currentBeginPosition)
		{
			var paragraphEndPos = stream.IndexOf("</p", currentBeginPosition, StringComparison.OrdinalIgnoreCase);
			var streamFinishedByParagraphEnd = stream.Substring(0, paragraphEndPos);
			var paragraphBeginPos = streamFinishedByParagraphEnd.LastIndexOf("<p", StringComparison.OrdinalIgnoreCase);
			var paragraphComplete = stream.Substring(paragraphBeginPos, paragraphEndPos-paragraphBeginPos);
			var indexOfOleObjectIdentifier = paragraphComplete.IndexOf("<o:OLEObject", StringComparison.OrdinalIgnoreCase);

			return indexOfOleObjectIdentifier != -1;
		}

		#endregion Private methods
	}
}
