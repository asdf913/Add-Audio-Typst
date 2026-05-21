package org.apache.pdfbox;

import java.awt.Component;
import java.awt.Container;
import java.awt.Desktop;
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
import java.io.Writer;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Member;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collection;
import java.util.EventObject;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiPredicate;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.IntConsumer;
import java.util.function.Predicate;
import java.util.stream.IntStream;
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
import javax.swing.table.TableModel;
import javax.swing.text.JTextComponent;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.function.FailableBiFunction;
import org.apache.commons.lang3.function.FailableConsumer;
import org.apache.commons.lang3.function.FailableFunction;
import org.apache.commons.lang3.function.FailableSupplier;
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

import com.j256.simplemagic.ContentInfo;
import com.j256.simplemagic.ContentInfoUtil;

import io.github.toolfactory.narcissus.Narcissus;
import net.miginfocom.swing.MigLayout;

public class AddAudioJPanel extends JPanel implements ActionListener {

	private static final long serialVersionUID = 2701307780245601666L;

	private static final String TYPST = "typst";

	private static final String VALUE = "value";

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

	@Note("Execute")
	private AbstractButton btnExecute;

	private AbstractButton btnBrowse;

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
		setEditable(tf, false);
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
			final Component component = getTableCellRendererComponent(tcr, jTable, value, isSelected, hasFocus, row,
					column);
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
		setEnabled(btnExecute, installed);
		//
		add(new JLabel("PDF"));
		//
		add(tfFilePdf = new JTextField(), growx);
		//
		add(btnBrowse = new JButton("Browse"), wrap);
		//
		setEnabled(btnBrowse, false);
		//
		new FailableStream<>(FieldUtils.getAllFieldsList(getClass()).stream().filter(f -> !isStatic(f))).forEach(f -> {
			//
			final Object object = Narcissus.getField(this, f);
			//
			setEditable(cast(JTextComponent.class, object), false);
			//
			addActionListener(cast(AbstractButton.class, object), this);
			//
		});
	}

	private static void setEnabled(final AbstractButton instance, final boolean enabled) {
		if (instance != null) {
			instance.setEnabled(enabled);
		}
	}

	private static Component getTableCellRendererComponent(final TableCellRenderer instnace, final JTable table,
			final Object value, boolean isSelected, boolean hasFocus, int row, int column) {
		return instnace != null
				? instnace.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column)
				: null;
	}

	private static void setEditable(final JTextComponent instance, final boolean editable) {
		if (instance != null) {
			instance.setEditable(editable);
		}
	}

	private static void addActionListener(final AbstractButton instance, final ActionListener actionListener) {
		if (instance != null) {
			instance.addActionListener(actionListener);
		}
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
			if (command != null && Narcissus.getObjectField(command, String.class.getDeclaredField(VALUE)) == null) {
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
		final Object source = getSource(evt);
		//
		if (Objects.equals(source, btnFileTemplate)) {
			//
			final JFileChooser jfc = new JFileChooser(".");
			//
			testAndRun(
					testAndGetAsBoolean(Boolean.logicalAnd(!GraphicsEnvironment.isHeadless(), !isTestMode()),
							() -> jfc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION),
					() -> setText(tfFileTemplate, getAbsolutePath(jfc.getSelectedFile())));
			//
			return;
			//
		} else if (Objects.equals(source, btnFileSpreadsheet)) {
			//
			final JFileChooser jfc = new JFileChooser(".");
			//
			File file = null;
			//
			testAndAccept(x -> and(x, AddAudioJPanel::exists, AddAudioJPanel::isFile),
					file = toFile(testAndApply(Objects::nonNull, getText(tfFileTemplate), Path::of, null)),
					x -> jfc.setCurrentDirectory(getParentFile(x)));
			//
			if (testAndGetAsBoolean(Boolean.logicalAnd(!GraphicsEnvironment.isHeadless(), !isTestMode()),
					() -> jfc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)) {
				//
				boolean isXlsx = false;
				//
				try {
					//
					setText(tfFileSpreadsheet,
							(isXlsx = isXlsx(Files.readAllBytes(Path.of(toURI(file = jfc.getSelectedFile())))))
									? getAbsolutePath(file)
									: null);
					//
				} catch (final IOException | SAXException | ParserConfigurationException e) {
					//
					setText(tfFileSpreadsheet, null);
					//
				} // try
					//
				forEach(IntStream.iterate(getRowCount(dtm) - 1, i -> i >= 0, i -> i - 1), i -> removeRow(dtm, i));
				//
				try (final Workbook wb = testAndApply(isXlsx, file, XSSFWorkbook::new, null)) {
					//
					forEach(values(createStringTextPositionEntryMap(
							testAndApply(x -> getNumberOfSheets(x) == 1, wb, x -> getSheetAt(x, 0), null),
							toFile(testAndApply(Objects::nonNull, getText(tfFileTemplate), Path::of, null)))),
							x -> addRow(dtm = ObjectUtils.getIfNull(dtm, DefaultTableModel::new), new Object[] { x }));
					//
				} catch (final Exception e) {
					//
					throw toRuntimeException(e);
					//
				} // try
					//
			} // if
				//
			return;
			//
		} else if (Objects.equals(source, btnBrowse)) {
			//
			testAndAccept(x -> and(x, AddAudioJPanel::exists, y -> isDirectory(getParentFile(y))),
					toFile(testAndApply(Objects::nonNull, getText(tfFilePdf), Path::of, null)),
					x -> open(testAndGet(Boolean.logicalAnd(!GraphicsEnvironment.isHeadless(), !isTestMode()),
							Desktop::getDesktop, null), getParentFile(x)),
					e -> {
						throw toRuntimeException(e);
					});
			//
			return;
			//
		} // if
			//
		actionPerformed(this, source);
		//
	}

	private static <T, E extends Throwable> void testAndAccept(final Predicate<T> predicate, final T value,
			final FailableConsumer<T, E> failableConsumer, final Consumer<Throwable> throwableConsumer) {
		//
		if (test(predicate, value)) {
			//
			try {
				//
				accept(failableConsumer, value);
				//
			} catch (final Throwable e) {
				//
				accept(throwableConsumer, e);
				//
			} // try
				//
		} // if
			//
	}

	private static <T, E extends Throwable> void accept(final FailableConsumer<T, E> instance, final T value) throws E {
		if (instance != null) {
			instance.accept(value);
		}
	}

	private static <T> void accept(final Consumer<T> instance, final T value) {
		if (instance != null) {
			instance.accept(value);
		}
	}

	private static boolean isDirectory(final File instance) {
		return instance != null && instance.getPath() != null && instance.isDirectory();
	}

	private static void open(final Desktop instance, final File file) throws IOException {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(instance)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), "peer")).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		if ((field == null || Narcissus.getField(instance, field) != null) && file != null && file.getPath() != null) {
			//
			instance.open(file);
			//
		} // if
			//
	}

	private static <T, R, E extends Throwable> R testAndApply(final boolean condition, final T value,
			final FailableFunction<T, R, E> functionTrue, final FailableFunction<T, R, E> functionFalse) throws E {
		return condition ? apply(functionTrue, value) : apply(functionFalse, value);
	}

	private static RuntimeException toRuntimeException(final Throwable instance) {
		return instance instanceof RuntimeException re ? re : new RuntimeException(instance);
	}

	private static void actionPerformed(final AddAudioJPanel instance, final Object source) {
		//
		if (instance == null || !Objects.equals(source, instance.btnExecute)) {
			//
			return;
			//
		} // if
			//
		setText(instance.tfFilePdf, null);
		//
		setEnabled(instance.btnBrowse, false);
		//
		Map<String, TextPositionEntry> map = null;
		//
		try (final Workbook wb = testAndApply(AddAudioJPanel::isFile,
				testAndApply(Objects::nonNull, getText(instance.tfFileSpreadsheet), File::new, null), XSSFWorkbook::new,
				null)) {
			//
			map = createStringTextPositionEntryMap(
					testAndApply(x -> getNumberOfSheets(x) == 1, wb, x -> getSheetAt(x, 0), null),
					toFile(testAndApply(Objects::nonNull, getText(instance.tfFileTemplate), Path::of, null)));
			//
		} catch (final Exception e) {
			//
			throw toRuntimeException(e);
			//
		} // try
			//
		final String outputPdf = String.join(".",
				StringUtils.substringBeforeLast(getText(instance.tfFileTemplate), "."), "pdf");
		//
		Process process = null;
		//
		PDDocument pdDocument = null;
		//
		try {
			//
			if (!isTestMode()
					&& (process = new ProcessBuilder(TYPST, "compile",
							StringUtils.defaultString(getText(instance.tfFileTemplate)), outputPdf).start()) != null
					&& process.waitFor() == 0) {
				// ;
				//
				final GetTextLocation pdfTextStripper = new GetTextLocation(map);
				//
				pdfTextStripper.setStartPage(1);
				//
				if ((pdDocument = Loader.loadPDF(Files.readAllBytes(Path.of(outputPdf)))) != null) {
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
			//
		try (final BufferedWriter writer = testAndApply(Objects::nonNull,
				testAndApply(Objects::nonNull,
						getOutputStream(process = testAndGet(!isTestMode(),
								() -> new ProcessBuilder(TYPST, "compile", "-", outputPdf).start(), null)),
						OutputStreamWriter::new, null),
				BufferedWriter::new, null)) {
			//
			write(writer,
					replace(testAndApply(AddAudioJPanel::isFile,
							testAndApply(Objects::nonNull, getText(instance.tfFileTemplate), File::new, null),
							x -> Files.readString(Path.of(toURI(x))), null), keySet(map), "\\u{25B6}"));
			//
		} catch (final IOException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
		try {
			//
			if (process != null && process.waitFor() == 0 && addPDAnnotations(map,
					pdDocument = Loader.loadPDF(Files.readAllBytes(Path.of(outputPdf))), getPage(pdDocument, 0))) {
				//
				final File file = toFile(Path.of(outputPdf));
				//
				save(pdDocument, file);
				//
				if (exists(getParentFile(file))) {
					//
					setEnabled(instance.btnBrowse, isDirectory(getParentFile(file)));
					//
				} // if
					//
				setText(instance.tfFilePdf, outputPdf);
				//
			} // if
				//
		} catch (final InterruptedException | IOException e) {
			//
			throw new RuntimeException(e);
			//
		} // try
			//
	}

	private static void save(final PDDocument instance, final File file) throws IOException {
		if (instance != null && file != null && exists(file) && isFile(file) && file.getPath() != null) {
			instance.save(file);
		}
	}

	private static boolean addPDAnnotations(final Map<String, TextPositionEntry> map, final PDDocument pdDocument,
			final PDPage pdPage) throws IOException {
		//
		if (iterator(entrySet(map)) != null) {
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
			TextPositionEntry textPositionEntry = null;
			//
			for (final Entry<String, TextPositionEntry> entry : entrySet(map)) {
				//
				if ((textPositionEntry = getValue(entry)) == null
						|| (textPosition = textPositionEntry.textPosition) == null) {
					//
					continue;
					//
				} // if
					//
				(pdComplexFileSpecification = new PDComplexFileSpecification())
						.setFile(getName(file = textPositionEntry.file));
				//
				final byte[] bs = testAndApply(Objects::nonNull, toPath(getAbsoluteFile(file)), Files::readAllBytes,
						null);
				//
				final ContentInfo ci = testAndApply(Objects::nonNull, bs, new ContentInfoUtil()::findMatch, null);
				//
				final StringBuilder sb = new StringBuilder(
						Integer.toString(Math.addExact(IterableUtils.size(getAnnotations(pdPage)), 1)));
				//
				if (ci != null && ci.getFileExtensions() != null && ci.getFileExtensions().length == 1) {
					//
					sb.append(".");
					//
					sb.append(ArrayUtils.get(ci.getFileExtensions(), 0));
					//
				} else {
					//
					sb.append(".wav");
					//
				} // if
					//
				pdComplexFileSpecification.setFile(Objects.toString(sb));
				//
				try (final InputStream is = testAndApply(Objects::nonNull, bs, ByteArrayInputStream::new, null)) {
					//
					if ((pdEmbeddedFile = testAndApply(Objects::nonNull, pdDocument, x -> new PDEmbeddedFile(x, is),
							null)) != null) {
						//
						pdEmbeddedFile.setSubtype(Objects.toString(ci != null ? ci.getMimeType() : null, "audio/wav"));
						//
					} // if
						//
					pdComplexFileSpecification.setEmbeddedFile(pdEmbeddedFile);
					//
				} // try
					//
				(pdAnnotationFileAttachment = new PDAnnotationFileAttachment()).setFile(pdComplexFileSpecification);
				//
				pdAnnotationFileAttachment.setRectangle(new PDRectangle(textPosition.getX(),
						(pdPage != null ? pdPage.getMediaBox().getHeight() : 0) - textPosition.getY(),
						getWidth(textPosition), textPosition.getHeight()));
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
			return true;
			//
		} // if
			//
		return false;
		//
	}

	private static float getWidth(final TextPosition instance) {
		//
		if (instance == null) {
			//
			return 0;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(instance)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), "textMatrix")).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		return field == null || Narcissus.getField(instance, field) != null ? instance.getWidth() : 0;
		//
	}

	private static Path toPath(final File instance) {
		return instance != null && instance.getPath() != null ? instance.toPath() : null;
	}

	private static File getAbsoluteFile(final File instance) {
		return instance != null && instance.getPath() != null ? instance.getAbsoluteFile() : null;
	}

	private static <T> void forEach(final Iterable<T> instance, final Consumer<T> consumer) {
		if (instance != null) {
			instance.forEach(consumer);
		}
	}

	private static Map<String, TextPositionEntry> createStringTextPositionEntryMap(final Iterable<Row> rows,
			final File file) {
		//
		Map<String, TextPositionEntry> map = null;
		//
		if (iterator(rows) != null) {
			//
			TextPositionEntry textPositionEntry = null;
			//
			Cell cell2 = null;
			//
			String s1, s2 = null;
			//
			Path path = null;
			//
			for (final Row row : rows) {
				//
				if ((cell2 = getCell(row, 2)) == null) {
					//
					continue;
					//
				} // if
					//
				path = null;
				//
				if (Boolean.logicalAnd((s1 = getAbsolutePath(getParentFile(file))) != null,
						(s2 = getStringCellValue(cell2)) != null)) {
					//
					path = Path.of(s1, s2);
					//
				} // if
					//
				(textPositionEntry = new TextPositionEntry()).file = toFile(path);
				//
				textPositionEntry.text = getStringCellValue(getCell(row, 1));
				//
				put(map = ObjectUtils.getIfNull(map, LinkedHashMap::new),
						textPositionEntry.marker = getStringCellValue(getCell(row, 0)), textPositionEntry);
				//
			} // for
				//
		} // if
			//
		return map;
		//
	}

	private static void testAndRun(final boolean condition, final Runnable runnable) {
		if (condition && runnable != null) {
			runnable.run();
		}
	}

	private static <T> boolean and(final T value, final Predicate<T> a, final Predicate<T> b) {
		return test(a, value) && test(b, value);
	}

	private static File getParentFile(final File instance) {
		return instance != null && instance.getPath() != null ? instance.getParentFile() : null;
	}

	private static <T> void testAndAccept(final Predicate<T> predicate, final T value, final Consumer<T> consumer) {
		if (test(predicate, value)) {
			accept(consumer, value);
		}
	}

	private static <K, V> void put(final Map<K, V> instance, final K key, final V value) {
		if (instance != null) {
			instance.put(key, value);
		}
	}

	private static <K> Set<K> keySet(final Map<K, ?> instance) {
		return instance != null ? instance.keySet() : null;
	}

	private static <V> Collection<V> values(final Map<?, V> instance) {
		return instance != null ? instance.values() : null;
	}

	private static <K, V> Collection<Entry<K, V>> entrySet(final Map<K, V> instance) {
		return instance != null ? instance.entrySet() : null;
	}

	private static Cell getCell(final Row instance, final int cellnum) {
		return instance != null ? instance.getCell(cellnum) : null;
	}

	private static String replace(final String string, final Iterable<String> ss, final String replacement) {
		//
		if (string == null) {
			//
			return null;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(string)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), VALUE)).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		if (field != null && Narcissus.getField(string, field) == null) {
			//
			return null;
			//
		} // if
			//
		String result = Objects.toString(new StringBuilder(string));
		//
		if (iterator(ss) != null) {
			//
			TextStringBuilder tsb = null;
			//
			for (final String s : ss) {
				//
				if ((tsb = ObjectUtils.getIfNull(tsb, TextStringBuilder::new)) == null) {
					//
					continue;
					//
				} // if
					//
				tsb.clear();
				//
				tsb.append(StringEscapeUtils.escapeJava(s));
				//
				tsb.append('}');
				//
				if (tsb.length() > 2) {
					//
					tsb.insert(2, '{');
					//
				} // if
					//
				result = replace(result, tsb, replacement);
				//
			} // for
				//
		} // if
			//
		return result;
		//
	}

	private static String replace(final String instance, final CharSequence target, final CharSequence replacement) {
		return instance != null && replacement != null ? instance.replace(target, replacement) : instance;
	}

	private static <T> Iterator<T> iterator(final Iterable<T> instance) {
		return instance != null ? instance.iterator() : null;
	}

	private static boolean testAndGetAsBoolean(final boolean condition, final BooleanSupplier supplier) {
		return condition && supplier != null && supplier.getAsBoolean();
	}

	private static Sheet getSheetAt(final Workbook instance, final int index) {
		return instance != null ? instance.getSheetAt(index) : null;
	}

	private static int getNumberOfSheets(final Workbook instance) {
		return instance != null ? instance.getNumberOfSheets() : 0;
	}

	private static void forEach(final IntStream instance, final IntConsumer consumer) {
		if (instance != null) {
			instance.forEach(consumer);
		}
	}

	private static int getRowCount(final TableModel instance) {
		return instance != null ? instance.getRowCount() : 0;
	}

	private static void addRow(final DefaultTableModel instance, final Object[] row) {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(instance)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), "dataVector")).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		if ((field == null || Narcissus.getField(instance, field) != null)) {
			//
			instance.addRow(row);
			//
		} // if
			//
	}

	private static void removeRow(final DefaultTableModel instance, final int row) {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(instance)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), "dataVector")).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		if ((field == null || Narcissus.getField(instance, field) != null) && getRowCount(instance) > row) {
			//
			instance.removeRow(row);
			//
		} // if
			//
	}

	private static Object getSource(final EventObject instance) {
		return instance != null ? instance.getSource() : null;
	}

	private static <T, E extends Throwable> T testAndGet(final boolean condition,
			final FailableSupplier<T, E> supplierTrue, final FailableSupplier<T, E> supplierFalse) throws E {
		return condition ? get(supplierTrue) : get(supplierFalse);
	}

	private static <T, E extends Throwable> T get(final FailableSupplier<T, E> instance) throws E {
		return instance != null ? instance.get() : null;
	}

	private static PDPage getPage(final PDDocument instance, final int pageIndex) {
		//
		if (instance == null) {
			//
			return null;
			//
		} // if
			//
		final Field field = testAndApply(x -> IterableUtils.size(x) == 1,
				FieldUtils.getAllFieldsList(getClass(instance)).stream()
						.filter(f -> f != null && Objects.equals(f.getName(), "document")).toList(),
				x -> IterableUtils.get(x, 0), null);
		//
		return (field == null || Narcissus.getField(instance, field) != null) && instance.getNumberOfPages() > pageIndex
				? instance.getPage(pageIndex)
				: null;
		//
	}

	private static URI toURI(final File instance) {
		return instance != null && instance.getPath() != null ? instance.toURI() : null;
	}

	private static void write(final Writer instance, final String string) throws IOException {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
		if (Arrays.stream(getClass(instance).getDeclaredMethods())
				.noneMatch(m -> m != null && Objects.equals(m.getName(), "write")
						&& Arrays.equals(m.getParameterTypes(), new Class<?>[] { String.class }))) {
			//
			final Method method = testAndApply(x -> IterableUtils.size(x) == 1,
					Arrays.stream(getClass(instance).getMethods())
							.filter(m -> m != null && Objects.equals(m.getName(), "write")
									&& Arrays.equals(m.getParameterTypes(), new Class<?>[] { String.class }))
							.toList(),
					x -> IterableUtils.get(x, 0), null);
			//
			if (method != null && Objects.equals(method.getDeclaringClass(), Writer.class) && string == null) {
				//
				return;
				//
			} // if
				//
		} // if
			//
		instance.write(string);
		//
	}

	private static <V> V getValue(final Entry<?, V> instance) {
		return instance != null ? instance.getValue() : null;
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
			if (Narcissus.getField(instance, Narcissus.findField(getClass(instance), "model")) == null
					|| (text != null && Narcissus.getField(text, Narcissus.findField(getClass(text), VALUE)) == null)) {
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

		private static void setTextPosition(final TextPositionEntry instance, final TextPosition textPosition) {
			if (instance != null) {
				instance.setTextPosition(textPosition);
			}
		}

		private void setTextPosition(final TextPosition textPosition) {
			this.textPosition = textPosition;
		}

	}

	public static void main(final String[] args) {
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
			if (iterator(textPositions) == null) {
				//
				return;
				//
			} // if
				//
			TextPositionEntry textPositionEntry = null;
			//
			for (final TextPosition textPosition : textPositions) {
				//
				if (textPosition == null
						|| iterator(keySet(map = ObjectUtils.getIfNull(map, LinkedHashMap::new))) == null) {
					//
					continue;
					//
				} // if
					//
				for (final String key : keySet(map)) {
					//
					if (Objects.equals(textPosition.getUnicode(), key)) {
						//
						if ((textPositionEntry = get(map, key)) == null) {
							//
							put(map, key, textPositionEntry = new TextPositionEntry());
							//
						} // if
							//
						TextPositionEntry.setTextPosition(textPositionEntry, textPosition);
						//
					} // if
						//
				} // for
					//
			} // for
				//
		}

		private static <V> V get(final Map<?, V> instance, final Object key) {
			return instance != null ? instance.get(key) : null;
		}

	}

}