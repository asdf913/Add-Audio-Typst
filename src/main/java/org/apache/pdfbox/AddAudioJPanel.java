package org.apache.pdfbox;

import java.awt.Component;
import java.awt.Container;
import java.awt.GraphicsEnvironment;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Member;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.function.BiPredicate;
import java.util.function.Predicate;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.swing.AbstractButton;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.WindowConstants;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.text.JTextComponent;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.function.FailableBiFunction;
import org.apache.commons.lang3.function.FailableFunction;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.stream.Streams.FailableStream;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.commons.text.TextStringBuilder;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.filespecification.PDComplexFileSpecification;
import org.apache.pdfbox.pdmodel.common.filespecification.PDEmbeddedFile;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationFileAttachment;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import org.zeroturnaround.zip.ZipUtil;

import io.github.toolfactory.narcissus.Narcissus;
import net.miginfocom.swing.MigLayout;

public class AddAudioJPanel extends JPanel implements ActionListener {

	private static final long serialVersionUID = 2701307780245601666L;

	private static final String TYPST = "typst";

	@Target(ElementType.FIELD)
	@Retention(RetentionPolicy.RUNTIME)
	private @interface Note {
		String value();
	}

	@Note("Template")
	private JTextComponent tfFileTemplate;

	@Note("Spreadsheet")
	private JTextComponent tfFileSpreadsheet;

	private JTextComponent tfFilePdf;

	@Note("Template")
	private AbstractButton btnFileTemplate;

	@Note("Spreadsheet")
	private AbstractButton btnFileSpreadsheet;

	private AbstractButton btnExecute;

	private DefaultTableModel dtm = null;

	private AddAudioJPanel() {
		//
		init();
		//
	}

	private void init() {
		//
		setLayout(new MigLayout());
		//
		try {
			//
			if (Narcissus.getObjectField(this, Container.class.getDeclaredField("component")) == null) {
				//
				return;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		add(new JLabel(TYPST), "span %1$s".formatted(1));
		//
		boolean installed = false;
		//
		try {
			//
			installed = exists(TYPST);
			//
		} catch (final IOException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		final JTextComponent tf = new JTextField(installed ? "Instlled" : "Not Installed");
		//
		final String growx = "growx";
		//
		final String wrap = "wrap";
		//
		add(tf, String.join(",", growx, wrap));
		//
		tf.setEditable(false);
		//
		add(new JLabel("Template"));
		//
		add(tfFileTemplate = new JTextField(), growx);
		//
		add(btnFileTemplate = new JButton("Select"), wrap);
		//
		add(new JLabel("Spreadsheet"));
		//
		add(tfFileSpreadsheet = new JTextField(), growx);
		//
		add(btnFileSpreadsheet = new JButton("Select"), wrap);
		//
		add(new JLabel());
		//
		final JTable jTable = new JTable(dtm = new DefaultTableModel(new Object[] { "Marker", "Text", "File" }, 0) {
			@Override
			public boolean isCellEditable(final int row, final int column) {
				return false;
			}
		});
		//
		add(new JScrollPane(jTable), wrap);
		//
		final TableCellRenderer tcr = jTable.getDefaultRenderer(Object.class);
		//
		jTable.setDefaultRenderer(Object.class, (table, value, isSelected, hasFocus, row, column) -> {
			//
			final TextPositionEntry textPositionEntry = cast(TextPositionEntry.class,
					dtm != null && dtm.getColumnCount() > 0 ? dtm.getValueAt(row, 0) : null);
			//
			final Component component = tcr != null
					? tcr.getTableCellRendererComponent(jTable, value, isSelected, hasFocus, row, column)
					: null;
			//
			final JLabel jLabel = cast(JLabel.class, component);
			//
			if (textPositionEntry != null) {
				//
				if (column == 0) {
					//
					setText(jLabel, textPositionEntry.marker);
					//
				} else if (column == 1) {
					//
					setText(jLabel, textPositionEntry.text);
					//
				} else if (column == 2) {
					//
					setText(jLabel, getName(textPositionEntry.file));
					//
				} // if
					//
			} // if
				//
			return ObjectUtils.getIfNull(jLabel, component);
			//
		});
		//
		add(new JLabel());
		//
		add(btnExecute = new JButton("Execute"), wrap);
		//
		btnExecute.setEnabled(installed);
		//
		add(new JLabel("PDF"));
		//
		add(tfFilePdf = new JTextField(), growx);
		//
		new FailableStream<>(FieldUtils.getAllFieldsList(getClass()).stream()
				.filter(f -> f != null && !Modifier.isStatic(f.getModifiers()))).forEach(f -> {
					//
					final Object object = Narcissus.getField(this, f);
					//
					final JTextComponent jtc = cast(JTextComponent.class, object);
					//
					if (jtc != null) {
						//
						jtc.setEditable(false);
						//
					} // if
						//
					final AbstractButton btn = cast(AbstractButton.class, object);
					//
					if (btn != null) {
						//
						btn.addActionListener(this);
						//
					} // if
						//
				});
	}

	private static void setText(final JLabel instance, final String text) {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		try {
			//
			if (Narcissus.getObjectField(instance, Component.class.getDeclaredField("objectLock")) == null) {
				//
				return;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		instance.setText(text);
		//
	}

	private static boolean exists(final String command) throws IOException {
		//
		return exists(getFile(command));
		//
	}

	private static boolean exists(final File instance) {
		//
		return instance != null && instance.getPath() != null && instance.exists();
		//
	}

	private static File getFile(final String command) throws IOException {
		//
		try {
			//
			if (command != null && Narcissus.getObjectField(command, String.class.getDeclaredField("value")) == null) {
				//
				return null;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		String whichOrWhere = "which";
		//
		if (Objects.equals(getName(getClass(FileSystems.getDefault())), "sun.nio.fs.WindowsFileSystem")) {
			//
			whichOrWhere = "where";
			//
		} // if
			//
		try (final InputStream is = getInputStream(
				start(testAndApply((a, b) -> b != null, whichOrWhere, command, ProcessBuilder::new, null)))) {
			//
			return testAndApply(Objects::nonNull,
					StringUtils.trim(testAndApply(Objects::nonNull, is,
							x -> new String(x != null ? x.readAllBytes() : null, StandardCharsets.UTF_8), null)),
					File::new, null);
			//
		} // try
			//
	}

	private static <T, U, R, E extends Throwable> R testAndApply(final BiPredicate<T, U> predicate, final T t,
			final U u, final FailableBiFunction<T, U, R, E> functionTrue,
			final FailableBiFunction<T, U, R, E> functionFalse) throws E {
		return test(predicate, t, u) ? apply(functionTrue, t, u) : apply(functionFalse, t, u);
	}

	private static <T, R, U, E extends Throwable> R apply(final FailableBiFunction<T, U, R, E> instance, final T t,
			final U u) throws E {
		return instance != null ? instance.apply(t, u) : null;
	}

	private static <T, U> boolean test(final BiPredicate<T, U> instance, final T t, final U u) {
		return instance != null && instance.test(t, u);
	}

	private static <T, R, E extends Throwable> R testAndApply(final Predicate<T> predicate, final T value,
			final FailableFunction<T, R, E> functionTrue, final FailableFunction<T, R, E> functionFalse) throws E {
		return test(predicate, value) ? apply(functionTrue, value) : apply(functionFalse, value);
	}

	private static <T> boolean test(final Predicate<T> instance, final T value) {
		return instance != null && instance.test(value);
	}

	private static <T, R, E extends Throwable> R apply(final FailableFunction<T, R, E> instance, final T value)
			throws E {
		return instance != null ? instance.apply(value) : null;
	}

	private static Process start(final ProcessBuilder instance) throws IOException {
		//
		if (instance == null) {
			//
			return null;
			//
		} // if
			//
		try {
			//
			if (Narcissus.getObjectField(instance, ProcessBuilder.class.getDeclaredField("command")) == null) {
				//
				return null;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		return instance.start();
		//
	}

	private static InputStream getInputStream(final Process instance) {
		return instance != null ? instance.getInputStream() : null;
	}

	private static String getName(final Class<?> instance) {
		return instance != null ? instance.getName() : null;
	}

	private static <T> T cast(final Class<T> clz, final Object instance) {
		return clz != null && clz.isInstance(instance) ? clz.cast(instance) : null;
	}

	@Override
	public void actionPerformed(final ActionEvent evt) {
		//
		final Object source = evt != null ? evt.getSource() : null;
		//
		if (Objects.equals(source, btnFileTemplate)) {
			//
			final JFileChooser jfc = new JFileChooser(".");
			//
			if (!GraphicsEnvironment.isHeadless() && !isTestMode()
					&& jfc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				//
				setText(tfFileTemplate, getAbsolutePath(jfc.getSelectedFile()));
				//
			} // if
				//
		} else if (Objects.equals(source, btnFileSpreadsheet)) {
			//
			final JFileChooser jfc = new JFileChooser(".");
			//
			File file = null;
			//
			if (exists(file = toFile(testAndApply(Objects::nonNull, getText(tfFileTemplate), Path::of, null)))
					&& isFile(file)) {
				//
				jfc.setCurrentDirectory(file.getParentFile());
				//
			} // if
				//
			if (!GraphicsEnvironment.isHeadless() && !isTestMode()
					&& jfc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				//
				boolean isXlsx = false;
				//
				try {
					//
					setText(tfFileSpreadsheet,
							(isXlsx = isXlsx(Files.readAllBytes(Path.of((file = jfc.getSelectedFile()).toURI()))))
									? getAbsolutePath(file)
									: null);
					//
				} catch (final IOException | SAXException | ParserConfigurationException e) {
					//
					setText(tfFileSpreadsheet, null);
					//
				} // try
					//
				for (int i = (dtm != null ? dtm.getRowCount() : 0) - 1; i >= 0; i--) {
					//
					dtm.removeRow(i);
					//
				} // for
					//
				if (isXlsx) {
					//
					try (final Workbook wb = new XSSFWorkbook(file)) {
						//
						final Sheet sheet = wb.getNumberOfSheets() == 1 ? wb.getSheetAt(0) : null;
						//
						if (sheet != null && sheet.iterator() != null) {
							//
							Cell cell2 = null;
							//
							TextPositionEntry textPositionEntry = null;
							//
							final File fileTemplate = toFile(Path.of(getText(tfFileTemplate)));
							//
							for (final Row row : sheet) {
								//
								if (row == null || (dtm = ObjectUtils.getIfNull(dtm, DefaultTableModel::new)) == null
										|| (cell2 = row.getCell(2)) == null) {
									//
									continue;
									//
								} // if
									//
								(textPositionEntry = new TextPositionEntry()).file = toFile(Path
										.of(fileTemplate.getParentFile().getAbsolutePath(), getStringCellValue(cell2)));
								//
								textPositionEntry.marker = getStringCellValue(row.getCell(0));
								//
								textPositionEntry.text = getStringCellValue(row.getCell(1));
								//
								dtm.addRow(new Object[] { textPositionEntry });
								//
							} // for
								//
						} // if
							//
					} catch (final InvalidFormatException | IOException e) {
						//
						throw e instanceof RuntimeException re ? re : new RuntimeException(e);
						//
					} // try
				} // if
					//
			} // if
				//
		} else if (Objects.equals(source, btnExecute)) {
			//
			setText(tfFilePdf, null);
			//
			Map<String, TextPositionEntry> map = null;
			//
			TextPositionEntry textPositionEntry = null;
			//
			try (final Workbook wb = testAndApply(AddAudioJPanel::isFile,
					testAndApply(Objects::nonNull, getText(tfFileSpreadsheet), File::new, null), XSSFWorkbook::new,
					null)) {
				//
				final Sheet sheet = wb != null && wb.getNumberOfSheets() == 1 ? wb.getSheetAt(0) : null;
				//
				if (sheet != null && sheet.iterator() != null) {
					//
					Cell cell2 = null;
					//
					final File fileTemplate = toFile(Path.of(getText(tfFileTemplate)));
					//
					for (final Row row : sheet) {
						//
						if (row == null || (map = ObjectUtils.getIfNull(map, LinkedHashMap::new)) == null
								|| (cell2 = row.getCell(2)) == null) {
							//
							continue;
							//
						} // if
							//
						(textPositionEntry = new TextPositionEntry()).file = toFile(
								Path.of(fileTemplate.getParentFile().getAbsolutePath(), getStringCellValue(cell2)));
						//
						textPositionEntry.text = getStringCellValue(row.getCell(1));
						//
						map.put(textPositionEntry.marker = getStringCellValue(row.getCell(0)), textPositionEntry);
						//
					} // for
						//
				} // if
					//
			} catch (final Exception e) {
				//
				throw e instanceof RuntimeException re ? re : new RuntimeException(e);
				//
			} // try
				//
			final String outputPdf = String.join(".", StringUtils.substringBeforeLast(getText(tfFileTemplate), "."),
					"pdf");
			//
			Process process = null;
			//
			PDDocument pdDocument = null;
			//
			PDPage pdPage = null;
			//
			try {
				//
				if (!isTestMode()
						&& (process = new ProcessBuilder(TYPST, "compile",
								StringUtils.defaultString(getText(tfFileTemplate)), outputPdf).start()) != null
						&& process.waitFor() == 0) {
					//
					pdPage = (pdDocument = Loader.loadPDF(Files.readAllBytes(Path.of(outputPdf)))) != null
							? pdDocument.getPage(0)
							: null;
					//
					final GetTextLocation pdfTextStripper = new GetTextLocation(map);
					//
					pdfTextStripper.setStartPage(1);
					//
					if (pdDocument != null) {
						//
						pdfTextStripper.setEndPage(pdDocument.getNumberOfPages());
						//
					} // if
						//
					pdfTextStripper.getText(pdDocument);
					//
				} // if
					//
			} catch (final IOException | InterruptedException e) {
				//
				throw new RuntimeException(e);
				//
			} // try
				// s
			try (final BufferedWriter writer = testAndApply(Objects::nonNull,
					testAndApply(Objects::nonNull,
							getOutputStream(!isTestMode()
									? process = new ProcessBuilder(TYPST, "compile", "-", outputPdf).start()
									: null),
							OutputStreamWriter::new, null),
					BufferedWriter::new, null)) {
				//
				String string = testAndApply(AddAudioJPanel::isFile,
						testAndApply(Objects::nonNull, getText(tfFileTemplate), File::new, null),
						x -> Files.readString(Path.of(x != null ? x.toURI() : null)), null);
				//
				TextStringBuilder tsb = null;
				//
				if (map != null && map.keySet() != null && map.keySet().iterator() != null) {
					//
					for (final String rowKey : map.keySet()) {
						//
						if ((tsb = ObjectUtils.getIfNull(tsb, TextStringBuilder::new)) == null) {
							//
							continue;
							//
						} // if
							//
						tsb.clear();
						//
						tsb.append(StringEscapeUtils.escapeJava(rowKey));
						//
						tsb.append('}');
						//
						tsb.insert(2, '{');
						//
						string = string.replace(tsb, "\\u{25B6}");
						//
					} // for
						//
				} // if
					//
				if (string != null) {
					//
					writer.write(string);
					//
				} // if
					//
			} catch (final IOException e) {
				//
				throw new RuntimeException(e);
				//
			} // try
				//
			try {
				//
				if (process != null && process.waitFor() == 0) {
					//
					pdPage = (pdDocument = Loader.loadPDF(Files.readAllBytes(Path.of(outputPdf)))) != null
							? pdDocument.getPage(0)
							: null;
					//
					if (map != null && map.entrySet() != null && map.entrySet().iterator() != null) {
						//
						PDComplexFileSpecification pdComplexFileSpecification = null;
						//
						PDEmbeddedFile pdEmbeddedFile = null;
						//
						PDAnnotationFileAttachment pdAnnotationFileAttachment = null;
						//
						TextPosition textPosition = null;
						//
						File file = null;
						//
						for (final Entry<String, TextPositionEntry> entry : map.entrySet()) {
							//
							if (entry == null || (textPositionEntry = entry.getValue()) == null
									|| (textPosition = textPositionEntry.textPosition) == null) {
								//
								continue;
								//
							} // if
								//
							(pdComplexFileSpecification = new PDComplexFileSpecification())
									.setFile(getName(file = textPositionEntry.file));
							//
							pdComplexFileSpecification
									.setFile(Math.addExact(IterableUtils.size(getAnnotations(pdPage)), 1) + ".wav");
							//
							try (final InputStream is = Files.newInputStream(file.getAbsoluteFile().toPath())) {
								//
								(pdEmbeddedFile = new PDEmbeddedFile(pdDocument, is)).setSubtype("audio/wav");
								//
								pdComplexFileSpecification.setEmbeddedFile(pdEmbeddedFile);
								//
							} // try
								//
							(pdAnnotationFileAttachment = new PDAnnotationFileAttachment())
									.setFile(pdComplexFileSpecification);
							//
							pdAnnotationFileAttachment.setRectangle(new PDRectangle(textPosition.getX(),
									pdPage.getMediaBox().getHeight() - textPosition.getY(), textPosition.getWidth(),
									textPosition.getHeight()));
							//
							pdAnnotationFileAttachment.setContents(textPositionEntry.text);
							//
							pdAnnotationFileAttachment.setConstantOpacity(0);
							//
							// 2. Mark as Locked (Bit 8) - prevents the annotation from being moved or
							// resized
							//
							// flags |= (1 << 7);
							//
							pdAnnotationFileAttachment
									.setAnnotationFlags(pdAnnotationFileAttachment.getAnnotationFlags() | (1 << 7));
							//
							add(getAnnotations(pdPage), pdAnnotationFileAttachment);
							//
						} // for
							//
						pdDocument.save(toFile(Path.of(outputPdf)));
						//
						setText(tfFilePdf, outputPdf);
					} // if
						//
				} // if
					//
			} catch (final InterruptedException | IOException e) {
				//
				throw new RuntimeException(e);
				//
			} // try
				//
		} // if
			//
	}

	private static OutputStream getOutputStream(final Process instance) {
		return instance != null ? instance.getOutputStream() : null;
	}

	private static File toFile(final Path instance) {
		return instance != null ? instance.toFile() : null;
	}

	private static List<PDAnnotation> getAnnotations(final PDPage instance) throws IOException {
		//
		if (instance == null) {
			//
			return null;
			//
		} // if
			//
		try {
			//
			if (Narcissus.getObjectField(instance, PDPage.class.getDeclaredField("page")) == null) {
				//
				return null;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		return instance.getAnnotations();
		//
	}

	private static boolean isFile(final File instance) {
		return instance != null && instance.getPath() != null && instance.isFile();
	}

	private static String getText(final JTextComponent instance) {
		//
		if (instance == null) {
			//
			return null;
			//
		} // if
			//
		try {
			//
			if (Narcissus.getField(instance, Narcissus.findField(getClass(instance), "model")) == null) {
				//
				return null;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		return instance.getText();
		//
	}

	public static boolean isXlsx(final byte[] bs) throws IOException, SAXException, ParserConfigurationException {
		//
		boolean contentTypeXmlFound = false;
		//
		try (final InputStream is = testAndApply(Objects::nonNull, bs, ByteArrayInputStream::new, null);
				final ZipInputStream zis = testAndApply(Objects::nonNull, is, ZipInputStream::new, null)) {
			//
			ZipEntry ze = null;
			//
			while ((ze = getNextEntry(zis)) != null) {
				//
				if (contentTypeXmlFound = Objects.equals("[Content_Types].xml", ze.getName())) {
					//
					break;
					//
				} // if
					//
			} // while
				//
		} // try
			//
		boolean isXlsx = false;
		//
		if (contentTypeXmlFound) {
			//
			try (final InputStream is = testAndApply(Objects::nonNull, bs, ByteArrayInputStream::new, null)) {
				//
				try (final InputStream bais = testAndApply(x -> x != null && x.length > 0,
						ZipUtil.unpackEntry(is, "[Content_Types].xml"), ByteArrayInputStream::new, null)) {
					//
					final NodeList childNodes = getChildNodes(getDocumentElement(
							bais != null ? parse(newDocumentBuilder(DocumentBuilderFactory.newDefaultInstance()), bais)
									: null));
					//
					for (int i = 0; i < getLength(childNodes); i++) {
						//
						if (Objects.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
								getTextContent(getNamedItem(getAttributes(item(childNodes, i)), "ContentType")))
								&& (isXlsx = true)) {
							//
							break;
							//
						} // if
							//
					} // for
						//
				} // try
					//
			} // try
				//
		} // if
			//
		return isXlsx;
		//
	}

	private static String getTextContent(final Node instance) {
		return instance != null ? instance.getTextContent() : null;
	}

	private static Node getNamedItem(final NamedNodeMap instance, final String name) {
		return instance != null ? instance.getNamedItem(name) : null;
	}

	private static NamedNodeMap getAttributes(final Node instance) {
		return instance != null ? instance.getAttributes() : null;
	}

	private static Node item(final NodeList instance, final int index) {
		return instance != null ? instance.item(index) : null;
	}

	private static int getLength(final NodeList instance) {
		return instance != null ? instance.getLength() : 0;
	}

	private static NodeList getChildNodes(final Node instance) {
		return instance != null ? instance.getChildNodes() : null;
	}

	private static Element getDocumentElement(final Document instance) {
		return instance != null ? instance.getDocumentElement() : null;
	}

	private static Document parse(final DocumentBuilder instance, final InputStream is)
			throws SAXException, IOException {
		return instance != null ? instance.parse(is) : null;
	}

	private static DocumentBuilder newDocumentBuilder(final DocumentBuilderFactory instance)
			throws ParserConfigurationException {
		return instance != null ? instance.newDocumentBuilder() : null;
	}

	private static ZipEntry getNextEntry(final ZipInputStream instance) {
		//
		if (instance == null) {
			//
			return null;
			//
		} // if
			//
		Object obj = null;
		//
		try {
			//
			final Method method = ZipInputStream.class.getDeclaredMethod("getNextEntry");
			//
			try {
				//
				if (Narcissus.getObjectField(instance, ZipInputStream.class.getDeclaredField("crc")) == null) {
					//
					return null;
					//
				} // if
					//
			} catch (final NoSuchFieldException e) {
				//
				throw new RuntimeException(e);
				//
			} // try
				//
			obj = Boolean.logicalOr(isStatic(method), instance != null) ? invoke(method, instance) : null;
			//
		} catch (final NoSuchMethodException | IllegalAccessException e) {
			//
			throw new RuntimeException(e);
			//
		} catch (final InvocationTargetException e) {
			//
			final Throwable targetException = e.getTargetException();
			//
			throw targetException instanceof RuntimeException re ? re : new RuntimeException(targetException);
			//
		} // try
			//
		return obj instanceof ZipEntry ze ? ze : null;
		//
	}

	private static Object invoke(final Method method, final Object instance, final Object... args)
			throws IllegalAccessException, InvocationTargetException {
		return method != null && method.getDeclaringClass() != null ? method.invoke(instance, args) : null;
	}

	private static boolean isStatic(final Member instance) {
		return instance != null && Modifier.isStatic(instance.getModifiers());
	}

	private static String getAbsolutePath(final File instance) {
		return instance != null && instance.getPath() != null ? instance.getAbsolutePath() : null;
	}

	private static void setText(final JTextComponent instance, final String text) {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		try {
			//
			if (Narcissus.getField(instance, Narcissus.findField(getClass(instance), "model")) == null || (text != null
					&& Narcissus.getField(text, Narcissus.findField(getClass(text), "value")) == null)) {
				//
				return;
				//
			} // if
				//
		} catch (final NoSuchFieldException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		instance.setText(text);
		//
	}

	private static Class<?> getClass(final Object instance) {
		return instance != null ? instance.getClass() : null;
	}

	private static class TextPositionEntry {

		private TextPosition textPosition;

		private String marker, text;

		private File file;

	}

	public static void main(final String[] args) throws InvalidFormatException, IOException, InterruptedException {
		//
		final JFrame jFrame = !GraphicsEnvironment.isHeadless() ? new JFrame() : null;
		//
		if (jFrame != null) {
			//
			jFrame.add(new AddAudioJPanel());
			//
			jFrame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
			//
			jFrame.pack();
			//
			if (!isTestMode()) {
				//
				jFrame.setVisible(true);
				//
			} // if
				//
		} // if
			//
	}

	private static boolean isTestMode() {
		try {
			return Class.forName("org.testng.annotations.Test") != null;
		} catch (final ClassNotFoundException e) {
			return false;
		}
	}

	private static <E> void add(final Collection<E> instance, final E item) {
		if (instance != null) {
			instance.add(item);
		}
	}

	private static String getStringCellValue(final Cell instance) {
		return instance != null ? instance.getStringCellValue() : null;
	}

	private static String getName(final File instance) {
		return instance != null && instance.getPath() != null ? instance.getName() : null;
	}

	private static class GetTextLocation extends PDFTextStripper {

		private Map<String, TextPositionEntry> map = null;

		private GetTextLocation(final Map<String, TextPositionEntry> map) {
			this.map = map;
		}

		@Override
		protected void writeString(final String string, final List<TextPosition> textPositions) throws IOException {
			//
			if (textPositions == null || textPositions.iterator() == null) {
				//
				return;
				//
			} // if
				//
			TextPositionEntry textPositionEntry = null;
			//
			for (final TextPosition textPosition : textPositions) {
				//
				if (textPosition == null || (map = ObjectUtils.getIfNull(map, LinkedHashMap::new)) == null
						|| map.keySet() == null || map.keySet().iterator() == null) {
					//
					continue;
					//
				} // if
					//
				for (final String key : map.keySet()) {
					//
					if (Objects.equals(textPosition.getUnicode(), key)) {
						//
						if ((textPositionEntry = map.get(key)) == null) {
							//
							map.put(key, textPositionEntry = new TextPositionEntry());
							//
						} // if
						if (textPositionEntry != null) {
							//
							textPositionEntry.textPosition = textPosition;
							//
						} // if
							//
					} // if
						//
				} // for
					//
			} // for
				//
		}
	}

}