/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package com.github.codeurjc.slidesconverter;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PowerPointToHTML {

	private static class Section {

		public String name;

		public int level;
		public int sectionNum;
		public int subsectionNum;
		public int startSlide;
		public Section parentSection;

		public Section(int level, int numInLevel, String name) {
			super();
			this.level = level;
			this.sectionNum = numInLevel;
			this.name = name;
		}
	}

	public static void main(String[] args) throws Exception {

		PowerPointToHTML converter = new PowerPointToHTML(
				Paths.get("../../1-Introducción.pptx"),
				Paths.get("../../1-Introducción.ppt"),
				Paths.get("../../out/Introduction.html"));

		converter.setMainTitle(1, 1, "Tema 1", "Introducción",
				"Desarrollo de aplicaciones web");
		
		converter.setTitleSlides(2, 14, 24, 62, 73, 85, 101, 112);
		
		converter.setImageSlide(7, 8, 9, 64, 82, 83, 84, 111, 29, 30, 32, 33,
				36, 40, 44, 50, 57, 58, 59, 60);
		
		converter.setNonSubtitleSlides(1, 26, 27, 34, 38, 42, 45, 54, 59, 77,
				113);

		converter.convert();

	}

	// Input parameters

	private int mainTitleFromSlide;
	private int mainTitleToSlide;
	private String mainTitleNumber;
	private String mainTitle;
	private String slidesContext;

	private int[] titleSlides = {};
	private int[] imageSlides = {};
	private int[] nonSubtitleSlides = {};

	// Conversion attributes

	private Section[] sectionsPerPage;

	private Path htmlFile;
	private Path pptxFile;
	private Path pptFile;

	private PrintWriter out;
	private double width;
	private double height;

	private void setMainTitle(int fromSlide, int toSlide, String number,
			String mainTitle, String context) {

		this.mainTitleFromSlide = fromSlide;
		this.mainTitleToSlide = toSlide;
		this.mainTitleNumber = number;
		this.mainTitle = mainTitle;
		this.slidesContext = context;
	}

	private void setNonSubtitleSlides(int... nonSubtitleSlides) {
		this.nonSubtitleSlides = nonSubtitleSlides;
	}

	private void setImageSlide(int... imageSlides) {
		this.imageSlides = imageSlides;
	}

	public void setTitleSlides(int... titleSlides) {
		this.titleSlides = titleSlides;
	}

	public PowerPointToHTML(Path pptxFile, Path pptFile, Path htmlFile) {

		this.pptxFile = pptxFile;
		this.pptFile = pptFile;
		this.htmlFile = htmlFile;
	}

	public void convert() throws IOException {

		InputStream fis = Files.newInputStream(pptxFile);
		XMLSlideShow pptx = new XMLSlideShow(fis);
		fis.close();

		InputStream is = Files.newInputStream(pptFile);
		HSLFSlideShow ppt = new HSLFSlideShow(is);
		is.close();

		width = pptx.getPageSize().getWidth();
		height = pptx.getPageSize().getHeight();

		out = new PrintWriter(Files.newOutputStream(htmlFile));

		out.println("<!DOCTYPE html>");
		out.println("<html><body>");

		out.println("<h1>" + this.mainTitleNumber + " " + mainTitle + "</h1>");
		out.println("<h2>" + this.slidesContext + "</h2>");

		List<Section> sections = calculateSections(pptx, ppt);

		generateTOC(sections);

		generateSlidesContent(pptx, ppt);

		pptx.close();
		ppt.close();
		out.close();

	}

	private void generateSlidesContent(XMLSlideShow pptx, HSLFSlideShow ppt)
			throws IOException {

		out.println("<h1>Slides</h1>");

		for (int numSlide = 0; numSlide < pptx.getSlides().size(); numSlide++) {

			System.out.println("Processing slide " + numSlide);

			XSLFSlide slideX = pptx.getSlides().get(numSlide);
			HSLFSlide slide = ppt.getSlides().get(numSlide);

			Section section = getSection(numSlide);

			if (section != null && section.startSlide == numSlide) {

				if (section.level == 0) {

					out.println("<h2>"
							+ escape(section.sectionNum + ". " + section.name)
							+ "</h2>");

				} else {

					out.println("<h3>" + escape(section.sectionNum + "."
							+ section.subsectionNum + ". " + section.name)
							+ "</h3>");
				}
			}

			generateSlideImage(slide);

			out.println("<h4>Slide (" + (numSlide + 1) + ")</h4>");

			out.println("<img src=" + getImageFileName(slideX.getSlideNumber())
					+ "><br>");

			for (XSLFShape shape : slideX.getShapes()) {

				if (shape instanceof XSLFTextShape) {

					XSLFTextShape tsh = (XSLFTextShape) shape;

					boolean first = true;
					boolean code = false;

					boolean subtitle = false;

					for (XSLFTextParagraph p : tsh) {

						String indent = "";

						int indentLevel = p.getIndentLevel();
						if (subtitle) {
							indentLevel++;
						}

						if (indentLevel > 0) {
							for (int j = 0; j < indentLevel; j++) {
								indent += "> ";
							}
						}

						StringBuilder sb = new StringBuilder();
						for (XSLFTextRun r : p) {
							sb.append(r.getRawText());
						}

						String text = sb.toString();

						// Avoid section or subsection name to appear in slide
						// contents
						if ((section != null && (text.equals(section.name)
								|| (section.parentSection != null
										&& section.parentSection.name
												.equals(text))))
								|| text.trim().isEmpty()) {
							continue;
						}

						if (first) {

							code = p.getDefaultFontFamily()
									.startsWith("Courier");

							if (code) {
								out.println(
										"<pre style='border:1px;border-style: solid;'>");
							}
						}

						if (!code) {
							out.println("<p>");
						}

						out(indent + text);

						if (!code) {
							out.println("</p>");
						}

						first = false;
						if (p.getTextAlign() == TextAlign.CENTER) {
							subtitle = true;
						}
					}

					if (code) {
						out.println("</pre>");
					}

				} else if (shape instanceof XSLFPictureShape) {

					XSLFPictureShape pShape = (XSLFPictureShape) shape;

					out.println("<p>Imagen: " + pShape.getShapeName() + "</p>");
				}
			}

			out.println("<hr>");
		}

		out.println("</body></html>");
	}

	private Section getSection(int slideNum) {

		return sectionsPerPage[slideNum];
	}

	private void generateTOC(List<Section> sections) {

		out.println("<h1>Contenido</h1>");
		out.println("<ul>");

		boolean lastInSubsection = false;

		for (Section section : sections) {
			if (section.level == 0) {

				if (lastInSubsection) {
					out.println("</ul></li>");
					lastInSubsection = false;
				}
				out.println("<li><strong>"
						+ escape(section.sectionNum + ". " + section.name + " ("
								+ (section.startSlide + 1) + ")")
						+ "</strong>");
			} else {

				if (!lastInSubsection) {
					out.println("<ul>");
				}

				out.println("<li>"
						+ escape(section.sectionNum + "."
								+ section.subsectionNum + ". " + section.name)
						+ " (" + (section.startSlide + 1) + ")" + "</li>");

				lastInSubsection = true;
			}
		}

		if (lastInSubsection) {
			out.println("</ul></li>");
			lastInSubsection = false;
		}

		out.println("</ul>");
	}

	private List<Section> calculateSections(XMLSlideShow pptx,
			HSLFSlideShow ppt) {

		sectionsPerPage = new Section[pptx.getSlides().size()];

		List<Section> sectionsList = new ArrayList<>();

		String lastSection = null;
		String lastSubsection = null;

		Section lastSectionObj = null;
		Section lastSubsectionObj = null;

		int numSections = 0;
		int numSubsectionsInSection = 0;

		for (int slideNum = 0; slideNum < pptx.getSlides().size(); slideNum++) {

			XSLFSlide slideX = pptx.getSlides().get(slideNum);
			HSLFSlide slide = ppt.getSlides().get(slideNum);

			String title = calculateTitle(slideX, slide);
			String subtitle = calculateSubtitle(slideX, slide);

			if (title != null && !title.trim().isEmpty()) {

				if (lastSection == null || !lastSection.equals(title)) {
					lastSection = title;
					numSections++;
					Section section = new Section(0, numSections, title);
					section.startSlide = slideNum;
					sectionsList.add(section);

					lastSubsection = null;
					lastSubsectionObj = null;

					numSubsectionsInSection = 0;

					this.sectionsPerPage[slideNum] = section;

					lastSectionObj = section;

					System.out.println("Section " + section.sectionNum + ". "
							+ section.name);
				}

			}

			if (subtitle != null && !subtitle.trim().isEmpty()) {

				if (lastSubsection == null
						|| !lastSubsection.equals(subtitle)) {
					lastSubsection = subtitle;
					numSubsectionsInSection++;

					Section section = new Section(1, numSections, subtitle);
					section.startSlide = slideNum;
					section.subsectionNum = numSubsectionsInSection;
					section.parentSection = lastSectionObj;
					sectionsList.add(section);

					this.sectionsPerPage[slideNum] = section;

					lastSubsectionObj = section;

					System.out.println("  Subsection " + section.sectionNum
							+ ". " + section.name);
				}

			}

			if (sectionsPerPage[slideNum] == null) {

				if (lastSubsectionObj != null) {
					this.sectionsPerPage[slideNum] = lastSubsectionObj;
				} else {
					this.sectionsPerPage[slideNum] = lastSectionObj;
				}
			}
		}

		return sectionsList;
	}

	private String calculateSubtitle(XSLFSlide slideX, HSLFSlide slide) {

		if (isContentSlide(slide.getSlideNumber())) {
			return null;
		}

		for (XSLFShape shape : slideX.getShapes()) {

			if (shape instanceof XSLFTextShape) {

				XSLFTextShape tsh = (XSLFTextShape) shape;

				Rectangle2D figure = getRelativeFigure(tsh);

				if (figure.getY() < 0.1) {
					continue;
				}

				for (XSLFTextParagraph p : tsh) {
					for (XSLFTextRun r : p) {
						return r.getRawText();
					}
				}

				return null;
			}
		}

		return null;
	}

	private boolean isContentSlide(int slideNum) {
		return (slideNum >= mainTitleFromSlide && slideNum <= mainTitleToSlide)
				|| ArrayUtils.contains(nonSubtitleSlides, slideNum)
				|| ArrayUtils.contains(titleSlides, slideNum)
				|| ArrayUtils.contains(imageSlides, slideNum);
	}

	private String escape(String text) {
		return StringEscapeUtils.escapeHtml4(text);
	}

	private String calculateTitle(XSLFSlide slideX, HSLFSlide slide) {

		String title = slide.getTitle();

		if (title != null) {
			return title;
		}

		title = slideX.getTitle();

		if (title != null) {
			return title;
		}

		boolean titleSlide = ArrayUtils.contains(titleSlides,
				slideX.getSlideNumber());

		for (XSLFShape shape : slideX.getShapes()) {

			if (shape instanceof XSLFTextShape) {

				XSLFTextShape tsh = (XSLFTextShape) shape;

				Rectangle2D figure = getRelativeFigure(tsh);

				if (titleSlide || figure.getY() < 0.1) {

					StringBuilder titleSB = new StringBuilder();

					for (XSLFTextParagraph p : tsh) {
						for (XSLFTextRun r : p) {
							titleSB.append(r.getRawText());
						}
					}

					title = titleSB.toString();

					if (!title.trim().isEmpty()) {
						return title;
					}
				}
			}
		}

		return null;
	}

	private Rectangle2D getRelativeFigure(XSLFTextShape tsh) {

		Rectangle2D anchor = tsh.getAnchor();

		return new Rectangle2D.Double(anchor.getX() * 100 / width,
				anchor.getY() * 100 / height, anchor.getWidth() * 100 / width,
				anchor.getHeight() * 100 / height);
	}

	private void out(String text) {
		out.print(StringEscapeUtils.escapeHtml4(text) + "\n");
	}

	public void generateSlideImage(HSLFSlide slide) throws IOException {

		float scale = 0.5f;

		int sWidth = (int) (width * scale);
		int sHeight = (int) (height * scale);

		String title = slide.getTitle();
		System.out.println("Rendering slide " + slide.getSlideNumber()
				+ (title == null ? "" : ": " + title));

		BufferedImage img = new BufferedImage(sWidth, sHeight,
				BufferedImage.TYPE_INT_RGB);
		Graphics2D graphics = img.createGraphics();
		graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
				RenderingHints.VALUE_ANTIALIAS_ON);
		graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
				RenderingHints.VALUE_RENDER_QUALITY);
		graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION,
				RenderingHints.VALUE_INTERPOLATION_BICUBIC);
		graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS,
				RenderingHints.VALUE_FRACTIONALMETRICS_ON);

		graphics.setPaint(Color.white);
		graphics.fill(new Rectangle2D.Float(0, 0, sWidth, sHeight));

		graphics.scale((double) sWidth / width, (double) sHeight / height);

		try {

			slide.draw(graphics);

			Path imagePath = htmlFile.getParent().resolve(
					Paths.get(getImageFileName(slide.getSlideNumber())));

			OutputStream out = Files.newOutputStream(imagePath);
			ImageIO.write(img, "png", out);
			out.close();

		} catch (Exception e) {
			System.err.println("Exception rendering slide "
					+ slide.getSlideNumber() + ": " + title);
		}
	}

	private String getImageFileName(int slideNumber) {
		return htmlFile.getFileName() + "-Slide" + slideNumber + ".png";
	}
}