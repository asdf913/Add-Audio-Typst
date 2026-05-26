package org.apache.pdfbox;

import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.lang.reflect.Array;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Member;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.Base64.Decoder;
import java.util.Calendar;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Predicate;
import java.util.stream.Stream;

import javax.swing.AbstractButton;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableModel;
import javax.swing.text.JTextComponent;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.bcel.generic.InstructionList;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.Strings;
import org.apache.commons.lang3.function.FailableBiConsumer;
import org.apache.commons.lang3.function.FailableBiFunction;
import org.apache.commons.lang3.function.FailableBiPredicate;
import org.apache.commons.lang3.function.FailableConsumer;
import org.apache.commons.lang3.function.FailableFunction;
import org.apache.commons.lang3.function.FailablePredicate;
import org.apache.commons.lang3.function.FailableSupplier;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.text.TextStringBuilder;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.common.filespecification.PDEmbeddedFile;
import org.apache.pdfbox.text.TextPosition;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.fasterxml.jackson.annotation.JsonAutoDetect.Visibility;
import com.google.common.base.Predicates;
import com.google.common.reflect.Reflection;
import com.j256.simplemagic.ContentInfo;
import com.j256.simplemagic.ContentInfoUtil;

import io.github.toolfactory.narcissus.Narcissus;
import javassist.util.proxy.MethodHandler;
import javassist.util.proxy.ProxyFactory;
import javassist.util.proxy.ProxyObject;
import tools.jackson.databind.json.JsonMapper;

class AddAudioJPanelTest {

	private static Method METHOD_EXISTS_FILE, METHOD_EXISTS_STRING, METHOD_GET_CLASS, METHOD_IS_FILE, METHOD_GET_NAME,
			METHOD_GET_ABSOLUTE_PATH, METHOD_TEST_AND_APPLY4, METHOD_TEST_AND_APPLY5, METHOD_INVOKE, METHOD_ADD,
			METHOD_CAST, METHOD_IS_STATIC, METHOD_IS_XLSX, METHOD_IS_XLS, METHOD_WRITE, METHOD_TO_URI, METHOD_GET_PAGE,
			METHOD_TEST_AND_GET, METHOD_REMOVE_ROW, METHOD_TEST_AND_GET_AS_BOOLEAN, METHOD_REPLACE, METHOD_ADD_ROW,
			METHOD_GET_PARENT_FILE, METHOD_TEST_AND_ACCEPT3, METHOD_TEST_AND_ACCEPT4_PREDICATE,
			METHOD_TEST_AND_ACCEPT4_FAILABLE_BI_PREDICATE, METHOD_AND, METHOD_CREATE_STRING_TEXT_POSITION_ENTRY_MAP,
			METHOD_ADD_PD_ANNOTATIONS, METHOD_TO_PATH, METHOD_GET_ABSOLUTE_FILE, METHOD_SAVE,
			METHOD_TO_RUNTIME_EXCEPTION, METHOD_OPEN, METHOD_IS_DIRECTORY, METHOD_SET_SUB_TYPE, METHOD_IIF,
			METHOD_GET_MEDIA_BOX, METHOD_GET_HEIGHT, METHOD_CLEAR, METHOD_APPEND_CHAR, METHOD_APPEND_INT,
			METHOD_APPEND_STRING, METHOD_GET_FILE_EXTENSION, METHOD_CONTAINS_KEY, METHOD_SET_MOD_DATE, METHOD_SET_SIZE,
			METHOD_CREATE_INPUT_STREAM_WORK_BOOK_FAILABLE_FUNCTION, METHOD_IS_SELECTED, METHOD_OR,
			METHOD_IS_VALID_TYPST_FILE, METHOD_GET_PAGES, METHOD_INDEX_OF, METHOD_CREATE_TYPST_FILE_FILTER,
			METHOD_TO_FILE = null;

	@BeforeSuite
	void beforeSuite() throws NoSuchMethodException {
		//
		final Class<?> clz = AddAudioJPanel.class;
		//
		(METHOD_EXISTS_FILE = clz.getDeclaredMethod("exists", File.class)).setAccessible(true);
		//
		(METHOD_EXISTS_STRING = clz.getDeclaredMethod("exists", String.class)).setAccessible(true);
		//
		(METHOD_GET_CLASS = clz.getDeclaredMethod("getClass", Object.class)).setAccessible(true);
		//
		(METHOD_IS_FILE = clz.getDeclaredMethod("isFile", File.class)).setAccessible(true);
		//
		(METHOD_GET_NAME = clz.getDeclaredMethod("getName", File.class)).setAccessible(true);
		//
		(METHOD_GET_ABSOLUTE_PATH = clz.getDeclaredMethod("getAbsolutePath", File.class)).setAccessible(true);
		//
		(METHOD_TEST_AND_APPLY5 = clz.getDeclaredMethod("testAndApply", FailableBiPredicate.class, Object.class,
				Object.class, FailableBiFunction.class, FailableBiFunction.class)).setAccessible(true);
		//
		(METHOD_INVOKE = clz.getDeclaredMethod("invoke", Method.class, Object.class, Object[].class))
				.setAccessible(true);
		//
		(METHOD_ADD = clz.getDeclaredMethod("add", Collection.class, Object.class)).setAccessible(true);
		//
		(METHOD_CAST = clz.getDeclaredMethod("cast", Class.class, Object.class)).setAccessible(true);
		//
		(METHOD_IS_STATIC = clz.getDeclaredMethod("isStatic", Member.class)).setAccessible(true);
		//
		(METHOD_IS_XLSX = clz.getDeclaredMethod("isXlsx", byte[].class)).setAccessible(true);
		//
		(METHOD_IS_XLS = clz.getDeclaredMethod("isXls", byte[].class)).setAccessible(true);
		//
		(METHOD_WRITE = clz.getDeclaredMethod("write", Writer.class, String.class)).setAccessible(true);
		//
		(METHOD_TO_URI = clz.getDeclaredMethod("toURI", File.class)).setAccessible(true);
		//
		(METHOD_GET_PAGE = clz.getDeclaredMethod("getPage", PDDocument.class, Integer.TYPE)).setAccessible(true);
		//
		(METHOD_GET_PAGE = clz.getDeclaredMethod("getPage", PDDocument.class, Integer.TYPE)).setAccessible(true);
		//
		(METHOD_TEST_AND_GET = clz.getDeclaredMethod("testAndGet", Boolean.TYPE, FailableSupplier.class,
				FailableSupplier.class)).setAccessible(true);
		//
		(METHOD_REMOVE_ROW = clz.getDeclaredMethod("removeRow", DefaultTableModel.class, Integer.TYPE))
				.setAccessible(true);
		//
		(METHOD_TEST_AND_GET_AS_BOOLEAN = clz.getDeclaredMethod("testAndGetAsBoolean", Boolean.TYPE,
				BooleanSupplier.class)).setAccessible(true);
		//
		(METHOD_REPLACE = clz.getDeclaredMethod("replace", String.class, Iterable.class, String.class))
				.setAccessible(true);
		//
		(METHOD_ADD_ROW = clz.getDeclaredMethod("addRow", DefaultTableModel.class, Object[].class)).setAccessible(true);
		//
		(METHOD_GET_PARENT_FILE = clz.getDeclaredMethod("getParentFile", File.class)).setAccessible(true);
		//
		(METHOD_TEST_AND_ACCEPT3 = clz.getDeclaredMethod("testAndAccept", Predicate.class, Object.class,
				FailableConsumer.class)).setAccessible(true);
		//
		(METHOD_TEST_AND_ACCEPT4_PREDICATE = clz.getDeclaredMethod("testAndAccept", Predicate.class, Object.class,
				FailableConsumer.class, Consumer.class)).setAccessible(true);
		//
		(METHOD_AND = clz.getDeclaredMethod("and", Object.class, Predicate.class, Predicate.class)).setAccessible(true);
		//
		(METHOD_CREATE_STRING_TEXT_POSITION_ENTRY_MAP = clz.getDeclaredMethod("createStringTextPositionEntryMap",
				Sheet.class, File.class)).setAccessible(true);
		//
		(METHOD_ADD_PD_ANNOTATIONS = clz.getDeclaredMethod("addPDAnnotations", Map.class, PDDocument.class,
				PDPage.class, Boolean.TYPE)).setAccessible(true);
		//
		(METHOD_TO_PATH = clz.getDeclaredMethod("toPath", File.class)).setAccessible(true);
		//
		(METHOD_GET_ABSOLUTE_FILE = clz.getDeclaredMethod("getAbsoluteFile", File.class)).setAccessible(true);
		//
		(METHOD_SAVE = clz.getDeclaredMethod("save", PDDocument.class, File.class)).setAccessible(true);
		//
		(METHOD_TO_RUNTIME_EXCEPTION = clz.getDeclaredMethod("toRuntimeException", Throwable.class))
				.setAccessible(true);
		//
		(METHOD_OPEN = clz.getDeclaredMethod("open", Desktop.class, File.class)).setAccessible(true);
		//
		(METHOD_IS_DIRECTORY = clz.getDeclaredMethod("isDirectory", File.class)).setAccessible(true);
		//
		(METHOD_SET_SUB_TYPE = clz.getDeclaredMethod("setSubtype", PDEmbeddedFile.class, String.class))
				.setAccessible(true);
		//
		(METHOD_IIF = clz.getDeclaredMethod("iif", Boolean.TYPE, Object.class, Object.class)).setAccessible(true);
		//
		(METHOD_GET_MEDIA_BOX = clz.getDeclaredMethod("getMediaBox", PDPage.class)).setAccessible(true);
		//
		(METHOD_GET_HEIGHT = clz.getDeclaredMethod("getHeight", PDRectangle.class)).setAccessible(true);
		//
		(METHOD_CLEAR = clz.getDeclaredMethod("clear", TextStringBuilder.class)).setAccessible(true);
		//
		(METHOD_APPEND_CHAR = clz.getDeclaredMethod("append", TextStringBuilder.class, Character.TYPE))
				.setAccessible(true);
		//
		(METHOD_APPEND_INT = clz.getDeclaredMethod("append", TextStringBuilder.class, Integer.TYPE))
				.setAccessible(true);
		//
		(METHOD_APPEND_STRING = clz.getDeclaredMethod("append", TextStringBuilder.class, String.class))
				.setAccessible(true);
		//
		(METHOD_GET_FILE_EXTENSION = clz.getDeclaredMethod("getFileExtension", ContentInfo.class)).setAccessible(true);
		//
		(METHOD_CONTAINS_KEY = clz.getDeclaredMethod("containsKey", Map.class, Object.class)).setAccessible(true);
		//
		(METHOD_SET_MOD_DATE = clz.getDeclaredMethod("setModDate", PDEmbeddedFile.class, Calendar.class))
				.setAccessible(true);
		//
		(METHOD_SET_SIZE = clz.getDeclaredMethod("setSize", PDEmbeddedFile.class, Integer.TYPE)).setAccessible(true);
		//
		(METHOD_CREATE_INPUT_STREAM_WORK_BOOK_FAILABLE_FUNCTION = clz
				.getDeclaredMethod("createInputStreamWorkbookFailableFunction", byte[].class)).setAccessible(true);
		//
		(METHOD_IS_SELECTED = clz.getDeclaredMethod("isSelected", AbstractButton.class)).setAccessible(true);
		//
		(METHOD_OR = clz.getDeclaredMethod("or", Object.class, FailablePredicate.class, FailablePredicate.class))
				.setAccessible(true);
		//
		(METHOD_IS_VALID_TYPST_FILE = clz.getDeclaredMethod("isValidTypstFile", File.class)).setAccessible(true);
		//
		(METHOD_GET_PAGES = clz.getDeclaredMethod("getPages", PDDocument.class)).setAccessible(true);
		//
		(METHOD_INDEX_OF = clz.getDeclaredMethod("indexOf", PDPageTree.class, PDPage.class)).setAccessible(true);
		//
		(METHOD_CREATE_TYPST_FILE_FILTER = clz.getDeclaredMethod("createTypstFileFilter")).setAccessible(true);
		//
		(METHOD_TO_FILE = clz.getDeclaredMethod("toFile", Path.class)).setAccessible(true);
		//
	}

	private static class IH implements InvocationHandler {

		private Integer length, modifiers, rowCount, numberOfSheets;

		private Boolean test, add, contains, getAsBoolean, containsKey, hasEntryCaseInsensitive;

		private Cell cell;

		private Throwable throwable;

		private Iterator<?> iterator;

		private String stringCellValue;

		@Override
		public Object invoke(final Object proxy, final Method method, final Object[] args) throws Throwable {
			//
			final String name = getName(method);
			//
			if (proxy instanceof FailableConsumer && Objects.equals(name, "accept")) {
				//
				if (throwable != null) {
					//
					throw throwable;
					//
				} // if
					//
				return null;
				//
			} // if
				//
			if (Objects.equals(getReturnType(method), Void.TYPE)) {
				//
				return null;
				//
			} else if (Objects.equals(name, "toString") && method != null && method.getParameterCount() == 0) {
				//
				return null;
				//
			} else if (proxy instanceof Iterable) {
				//
				if (Objects.equals(name, "iterator")) {
					//
					return iterator;
					//
				} else if (Objects.equals(name, "spliterator")) {
					//
					return null;
					//
				} // if
					//
			} else if (proxy instanceof Stream) {
				//
				if (Objects.equals(name, "filter")) {
					//
					return proxy;
					//
				} else if (Objects.equals(name, "toList")) {
					//
					return null;
					//
				} // if
					//
			} // if
				//
			if (proxy instanceof NodeList) {
				//
				if (Objects.equals(name, "getLength")) {
					//
					return length;
					//
				} else if (Objects.equals(name, "item")) {
					//
					return null;
					//
				} // if
					//
			} else if (Boolean.logicalAnd(Objects.equals(name, "test"),
					Stream.of(Predicate.class, FailableBiPredicate.class, FailablePredicate.class)
							.anyMatch(x -> x != null && proxy != null && x.isAssignableFrom(proxy.getClass())))) {
				//
				return test;
				//
			} else if (proxy instanceof Collection) {
				//
				if (Objects.equals(name, "add")) {
					//
					return add;
					//
				} else if (Objects.equals(name, "contains")) {
					//
					return contains;
					//
				} // if
					//
			} else if (Boolean.logicalAnd(Objects.equals(name, "apply"),
					Boolean.logicalOr(proxy instanceof FailableFunction, proxy instanceof FailableBiFunction))) {
				//
				return null;
				//
			} else if (proxy instanceof Member && Objects.equals(name, "getModifiers")) {
				//
				return modifiers;
				//
			} else if (proxy instanceof Node
					&& contains(Arrays.asList("getAttributes", "getTextContent", "getChildNodes"), name)) {
				//
				return null;
				//
			} else if (proxy instanceof Document && Objects.equals(name, "getDocumentElement")) {
				//
				return null;
				//
			} else if (proxy instanceof Cell && Objects.equals(name, "getStringCellValue")) {
				//
				return stringCellValue;
				//
			} else if (proxy instanceof NamedNodeMap && Objects.equals(name, "getNamedItem")) {
				//
				return null;
				//
			} else if (proxy instanceof Path && Objects.equals(name, "toFile")) {
				//
				return null;
				//
			} else if (proxy instanceof TableCellRenderer && Objects.equals(name, "getTableCellRendererComponent")) {
				//
				return null;
				//
			} else if (proxy instanceof Entry && Objects.equals(name, "getValue")) {
				//
				return null;
				//
			} else if (proxy instanceof FailableSupplier && Objects.equals(name, "get")) {
				//
				return null;
				//
			} else if (proxy instanceof TableModel && Objects.equals(name, "getRowCount")) {
				//
				return rowCount;
				//
			} else if (proxy instanceof Workbook) {
				//
				if (Objects.equals(name, "getSheetAt")) {
					//
					return null;
					//
				} else if (Objects.equals(name, "getNumberOfSheets")) {
					//
					return numberOfSheets;
					//
				} // if
					//
			} else if (proxy instanceof BooleanSupplier && Objects.equals(name, "getAsBoolean")) {
				//
				return getAsBoolean;
				//
			} else if (proxy instanceof Sheet && Objects.equals(name, "getSheetName")) {
				//
				return null;
				//
			} else if (proxy instanceof Row && Objects.equals(name, "getCell")) {
				//
				return cell;
				//
			} else if (proxy instanceof Map) {
				//
				if (contains(List.of("keySet", "put", "values", "entrySet", "get"), name)) {
					//
					return null;
					//
				} else if (Objects.equals(name, "containsKey")) {
					//
					return containsKey;
					//
				} // if
					//
			} else if (proxy instanceof Member && Objects.equals(name, "getName")) {
				//
				return null;
				//
			} else if (proxy instanceof DirectoryEntry) {
				//
				if (Objects.equals(name, "getEntryCaseInsensitive")) {
					//
					return null;
					//
				} else if (Objects.equals(name, "hasEntryCaseInsensitive")) {
					//
					return hasEntryCaseInsensitive;
					//
				} // if
					//
				return null;
				//
			} // if
				//
			throw new Throwable(name);
			//
		}

	}

	private IH ih = null;

	private AddAudioJPanel instance = null;

	private Decoder decoder = null;

	@BeforeMethod
	void beforeMethod() throws Throwable {
		//
		ih = new IH();
		//
		instance = cast(AddAudioJPanel.class, Narcissus.allocateInstance(AddAudioJPanel.class));
		//
		decoder = Base64.getDecoder();
		//
	}

	@Test
	void testNull() throws Throwable {
		//
		final Method[] ms = AddAudioJPanel.class.getDeclaredMethods();
		//
		Method m = null;
		//
		Class<?>[] parameterTypes = null;
		//
		Class<?> parameterType = null;
		//
		Object result = null;
		//
		String toString = null;
		//
		Object[] os = null;
		//
		Collection<Object> collection = null;
		//
		for (int i = 0; ms != null && i < ms.length; i++) {
			//
			if ((m = ArrayUtils.get(ms, i)) == null || m.isSynthetic()
					|| (parameterTypes = m.getParameterTypes()) == null) {
				//
				continue;
				//
			} // if
				//
			clear(collection = ObjectUtils.getIfNull(collection, ArrayList::new));
			//
			for (int j = 0; j < parameterTypes.length; j++) {
				//
				if (Objects.equals(parameterType = ArrayUtils.get(parameterTypes, j), Integer.TYPE)) {
					//
					add(collection, Integer.valueOf(0));
					//
				} else if (Objects.equals(parameterType, Long.TYPE)) {
					//
					add(collection, Long.valueOf(0));
					//
				} else if (Objects.equals(parameterType, Boolean.TYPE)) {
					//
					add(collection, Boolean.FALSE);
					//
				} else if (Objects.equals(parameterType, Character.TYPE)) {
					//
					add(collection, Character.valueOf(' '));
					//
				} else if (Objects.equals(parameterType, byte[].class)) {
					//
					add(collection, new byte[] { 0 });
					//
				} else {
					//
					add(collection, null);
					//
				} // if
					//
			} // for
				//
			os = toArray(collection);
			//
			toString = Objects.toString(m);
			//
			if (Modifier.isStatic(m.getModifiers())) {
				//
				result = Narcissus.invokeStaticMethod(m, os);
				//
			} else {
				//
				result = Narcissus
						.invokeMethod(
								instance = ObjectUtils.getIfNull(instance,
										() -> (AddAudioJPanel) Narcissus.allocateInstance(AddAudioJPanel.class)),
								m, os);
				//
			} // if
				//
			if (contains(Arrays.asList(Integer.TYPE, Boolean.TYPE, Float.TYPE, Long.TYPE), getReturnType(m))
					|| Boolean.logicalAnd(Objects.equals(getName(m), "toRuntimeException"),
							Arrays.equals(parameterTypes, new Class<?>[] { Throwable.class }))
					|| Boolean.logicalAnd(Objects.equals(getName(m), "createTypstFileFilter"),
							Arrays.equals(parameterTypes, new Class<?>[] {}))) {
				//
				Assert.assertNotNull(result, toString);
				//
			} else {
				//
				Assert.assertNull(result, toString);
				//
			} // if
				//
		} // for
			//
	}

	@Test
	void testNotNull() throws Throwable {
		//
		final Method[] ms = AddAudioJPanel.class.getDeclaredMethods();
		//
		Method m = null;
		//
		Class<?>[] parameterTypes = null;
		//
		Class<?> parameterType = null;
		//
		Object object, result = null;
		//
		String name, toString = null;
		//
		Object[] os = null;
		//
		Collection<Object> collection = null;
		//
		ProxyFactory proxyFactory = null;
		//
		for (int i = 0; ms != null && i < ms.length; i++) {
			//
			if ((m = ArrayUtils.get(ms, i)) == null || m.isSynthetic()
					|| (parameterTypes = m.getParameterTypes()) == null
					|| Boolean.logicalAnd(Objects.equals(name = getName(m), "main"),
							Arrays.equals(parameterTypes, new Class<?>[] { String[].class }))
					|| Boolean.logicalAnd(Objects.equals(name, "parse"),
							Arrays.equals(parameterTypes, new Class<?>[] { DocumentBuilder.class, InputStream.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getInputStream"),
							Arrays.equals(parameterTypes, new Class<?>[] { Process.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "replace"),
							Arrays.equals(parameterTypes,
									new Class<?>[] { String.class, CharSequence.class, CharSequence.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "append"),
							Arrays.equals(parameterTypes, new Class<?>[] { TextStringBuilder.class, String.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "append"),
							Arrays.equals(parameterTypes, new Class<?>[] { TextStringBuilder.class, Integer.TYPE }))
					|| Boolean.logicalAnd(Objects.equals(name, "append"),
							Arrays.equals(parameterTypes, new Class<?>[] { TextStringBuilder.class, Character.TYPE }))
					|| Boolean.logicalAnd(Objects.equals(name, "clear"),
							Arrays.equals(parameterTypes, new Class<?>[] { TextStringBuilder.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getDeclaredMethod"),
							Arrays.equals(parameterTypes, new Class<?>[] { Class.class, String.class, Class[].class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getEntryNames"),
							Arrays.equals(parameterTypes, new Class<?>[] { Class.class }))) {
				//
				continue;
				//
			} // if
				//
			clear(collection = ObjectUtils.getIfNull(collection, ArrayList::new));
			//
			for (int j = 0; j < parameterTypes.length; j++) {
				//
				if (Objects.equals(parameterType = ArrayUtils.get(parameterTypes, j), Integer.TYPE)) {
					//
					add(collection, Integer.valueOf(0));
					//
				} else if (Objects.equals(parameterType, Long.TYPE)) {
					//
					add(collection, Long.valueOf(0));
					//
				} else if (Objects.equals(parameterType, Boolean.TYPE)) {
					//
					add(collection, Boolean.FALSE);
					//
				} else if (Objects.equals(parameterType, Character.TYPE)) {
					//
					add(collection, Character.valueOf(' '));
					//
				} else if (Objects.equals(parameterType, Class.class)) {
					//
					add(collection, Object.class);
					//
				} else if (Objects.equals(parameterType, JTextComponent.class)) {
					//
					add(collection, new JTextField());
					//
				} else if (Objects.equals(parameterType, DocumentBuilderFactory.class)) {
					//
					add(collection, DocumentBuilderFactory.newDefaultInstance());
					//
				} else if (Objects.equals(parameterType, byte[].class)) {
					//
					add(collection, new byte[] { 0 });
					//
				} else if (contains(List.of(JComponent.class, AbstractButton.class), parameterType)) {
					//
					add(collection, new JButton());
					//
				} else if (Objects.equals(parameterType, Strings.class)) {
					//
					add(collection, Strings.CS);
					//
				} else if (Objects.equals(parameterType, Process.class) || Objects.equals(parameterType, Writer.class)
						|| Objects.equals(parameterType, Calendar.class)
						|| Objects.equals(parameterType, InputStream.class)) {
					//
					(proxyFactory = new ProxyFactory()).setSuperclass(parameterType);
					//
					if ((object = newInstance(
							getDeclaredConstructor(proxyFactory.createClass()))) instanceof ProxyObject) {
						//
						((ProxyObject) object).setHandler(new MethodHandler() {

							@Override
							public Object invoke(final Object self, final Method thisMethod, final Method proceed,
									final Object[] args) throws Throwable {
								//
								if (Objects.equals(getReturnType(thisMethod), Void.TYPE)) {
									//
									return null;
									//
								} // if
									//
								final String name = getName(thisMethod);
								//
								if (Boolean.logicalOr(
										Boolean.logicalAnd(self instanceof Process,
												contains(List.of("getOutputStream", "getErrorStream"), name)),
										Boolean.logicalAnd(self instanceof InputStream,
												Objects.equals(name, "readAllBytes")))) {
									//
									return null;
									//
								} // if
									//
								throw new Throwable(name);
								//
							}
						});
						//
					} // if
						//
					add(collection, object);
					//
				} else if (parameterType != null && parameterType.isArray()) {
					//
					add(collection, Array.newInstance(parameterType, 0));
					//
				} else if (parameterType.isInterface()) {
					//
					FieldUtils.getAllFieldsList(getClass(ih = ObjectUtils.getIfNull(ih, IH::new))).forEach(f -> {
						//
						final Class<?> type = f != null ? f.getType() : null;
						//
						if (Objects.equals(type, Integer.class)) {
							//
							Narcissus.setField(ih, f, Integer.valueOf(0));
							//
						} else if (Objects.equals(type, Boolean.class)) {
							//
							Narcissus.setField(ih, f, Boolean.FALSE);
							//
						} // if
							//
					});
					//
					add(collection, Reflection.newProxy(parameterType, ih));
					//
				} else {
					//
					add(collection, Narcissus.allocateInstance(parameterType));
					//
				} // if
					//
			} // for
				//
			os = toArray(collection);
			//
			toString = Objects.toString(m);
			//
			if (Modifier.isStatic(m.getModifiers())) {
				//
				result = Narcissus.invokeStaticMethod(m, os);
				//
			} else {
				//
				result = Narcissus
						.invokeMethod(
								instance = ObjectUtils.getIfNull(instance,
										() -> (AddAudioJPanel) Narcissus.allocateInstance(AddAudioJPanel.class)),
								m, os);
				//
			} // if
				//
			if (contains(Arrays.asList(Integer.TYPE, Boolean.TYPE, Float.TYPE), getReturnType(m))
					|| Boolean.logicalAnd(Objects.equals(name, "getName"),
							Arrays.equals(parameterTypes, new Class<?>[] { Class.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getClass"),
							Arrays.equals(parameterTypes, new Class<?>[] { Object.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "cast"),
							Arrays.equals(parameterTypes, new Class<?>[] { Class.class, Object.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getText"),
							Arrays.equals(parameterTypes, new Class<?>[] { JTextComponent.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "newDocumentBuilder"),
							Arrays.equals(parameterTypes, new Class<?>[] { DocumentBuilderFactory.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "toRuntimeException"),
							Arrays.equals(parameterTypes, new Class<?>[] { Throwable.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "iif"),
							Arrays.equals(parameterTypes, new Class<?>[] { Boolean.TYPE, Object.class, Object.class }))
					|| Boolean.logicalAnd(Objects.equals(name, "getInstructions"),
							Arrays.equals(parameterTypes, new Class<?>[] { InstructionList.class }))
					|| Boolean.logicalAnd(Objects.equals(getName(m), "createTypstFileFilter"),
							Arrays.equals(parameterTypes, new Class<?>[] {}))) {
				//
				Assert.assertNotNull(result, toString);
				//
			} else if (Boolean.logicalAnd(Objects.equals(name, "filter"),
					Arrays.equals(parameterTypes, new Class<?>[] { Stream.class, Predicate.class })) && os != null
					&& os.length > 0) {
				//
				Assert.assertSame(result, ArrayUtils.get(os, 0));
				//
			} else {
				//
				Assert.assertNull(result, toString);
				//
			} // if
				//
		} // for
			//
	}

	private static <T> Constructor<T> getDeclaredConstructor(final Class<T> instance, final Class<?>... parameterTypes)
			throws NoSuchMethodException {
		return instance != null ? instance.getDeclaredConstructor(parameterTypes) : null;
	}

	private static <T> T newInstance(final Constructor<T> instance, final Object... args)
			throws InstantiationException, IllegalAccessException, InvocationTargetException {
		return instance != null ? instance.newInstance(args) : null;
	}

	private static String getName(final Member instance) {
		return instance != null ? instance.getName() : null;
	}

	private static <E> void add(final Collection<E> items, final E item)
			throws IllegalAccessException, InvocationTargetException {
		//
		invoke(METHOD_ADD, null, items, item);
		//
	}

	private static void clear(final Collection<?> instance) {
		if (instance != null) {
			instance.clear();
		}
	}

	private static boolean contains(final Collection<?> items, final Object item) {
		return items != null && items.contains(item);
	}

	private static Class<?> getReturnType(final Method instance) {
		return instance != null ? instance.getReturnType() : null;
	}

	private static Object[] toArray(final Collection<?> instance) {
		return instance != null ? instance.toArray() : null;
	}

	private static Object invoke(final Method method, final Object instance, final Object... args)
			throws IllegalAccessException, InvocationTargetException {
		return METHOD_INVOKE != null ? METHOD_INVOKE.invoke(null, method, instance, args) : null;
	}

	@Test
	void testExists() throws Throwable {
		//
		Assert.assertEquals(invoke(METHOD_EXISTS_FILE, null, toFile(Path.of("."))), Boolean.TRUE);
		//
		Assert.assertEquals(invoke(METHOD_EXISTS_FILE, null, toFile(Path.of("1"))), Boolean.FALSE);
		//
	}

	private static File toFile(final Path instance) throws Throwable {
		try {
			final Object obj = invoke(METHOD_TO_FILE, null, instance);
			if (obj == null) {
				return null;
			} else if (obj instanceof File) {
				return (File) obj;
			} // if
			throw new Throwable(Objects.toString(getClass(obj)));
		} catch (final InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	private static Class<?> getClass(final Object instance) throws Throwable {
		try {
			final Object obj = invoke(METHOD_GET_CLASS, null, instance);
			if (obj == null) {
				return null;
			} else if (obj instanceof Class) {
				return (Class<?>) obj;
			} // if
			throw new Throwable(Objects.toString(getClass(obj)));
		} catch (final InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	@Test
	void testIsFile() throws Throwable {
		//
		Assert.assertEquals(invoke(METHOD_IS_FILE, null, toFile(Path.of("pom.xml"))), Boolean.TRUE);
		//
		Assert.assertEquals(invoke(METHOD_IS_FILE, null, toFile(Path.of("."))), Boolean.FALSE);
		//
	}

	@Test
	void testGetName() throws Throwable {
		//
		final String name = "pom.xml";
		//
		Assert.assertEquals(name, invoke(METHOD_GET_NAME, null, toFile(Path.of(name))));
		//
	}

	@Test
	void testGetAbsolutePath() throws Throwable {
		//
		Assert.assertNotNull(getAbsolutePath(toFile(Path.of("pom.xml"))));
		//
	}

	private static String getAbsolutePath(final File instance) throws Throwable {
		try {
			final Object obj = invoke(METHOD_GET_ABSOLUTE_PATH, null, instance);
			if (obj == null) {
				return null;
			} else if (obj instanceof String) {
				return (String) obj;
			} // if
			throw new Throwable(Objects.toString(getClass(obj)));
		} catch (final InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	@Test
	void testTestAndApply() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNull(invoke(METHOD_TEST_AND_APPLY4, null, Boolean.TRUE, null, null, null));
		//
		if ((ih = ObjectUtils.getIfNull(ih, IH::new)) != null) {
			//
			ih.test = Boolean.TRUE;
			//
		} // if
			//
		Assert.assertNull(invoke(METHOD_TEST_AND_APPLY5, null, Reflection.newProxy(FailableBiPredicate.class, ih), null,
				null, null, null));
		//
	}

	@Test
	void testCast() throws Throwable {
		//
		Assert.assertNull(cast(Object.class, null));
		//
	}

	private static <T> T cast(final Class<T> clz, final Object instance) throws Throwable {
		try {
			return (T) invoke(METHOD_CAST, null, clz, instance);
		} catch (final InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	@Test
	void testIsStatic() throws IllegalAccessException, InvocationTargetException, NoSuchFieldException {
		//
		Assert.assertEquals(invoke(METHOD_IS_STATIC, null, Boolean.class.getDeclaredField("TRUE")), Boolean.TRUE);
		//
	}

	@Test
	void testIsXlsx() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertEquals(invoke(METHOD_IS_XLSX, null, decode(decoder,
				"UEsDBBQACAgIANAAslwAAAAAAAAAAAAAAAAaAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHOtkcuqwjAQhvc+RZi9TasgB2nqRgS3og8Q0umFtknIjLe3NyoeFeRwFq6Gfy7f/5Pki9PQiwMGap1VkCUpCLTGla2tFey2q/EPLIpRvsFec1yhpvUk4o0lBQ2zn0tJpsFBU+I82jipXBg0Rxlq6bXpdI1ykqYzGV4ZULwxxbpUENZlBmJ79vgftquq1uDSmf2Alj9YSI63GIE61MgKbvLezJIIA/k5w+SbGYjPPdIzxF3/ZT/9pv3RhY4aRH4m+G3FcNfyeItRLt9+ubgAUEsHCPvN1gfNAAAAHAIAAFBLAwQUAAgICADQALJcAAAAAAAAAAAAAAAADwAAAHhsL3dvcmtib29rLnhtbKWUTXPaMBCG7/0VHt3BlgGXMDGZlIRJZvo1SZqcZXsdq5Elj7QEaKf/vWsZQzL0kEkPoI9dPXpXeuXTs02tgmewThqdMj6MWAA6N4XUjyn7cbccTFngUOhCKKMhZVtw7Gz+4XRt7FNmzFNA67VLWYXYzMLQ5RXUwg1NA5oipbG1QBrax9A1FkThKgCsVRhHURLWQmrWEWb2LQxTljKHC5OvatDYQSwogaTeVbJxPa3O34KrhX1aNYPc1A0hMqkkbj20x2z45IhTy9waZ0oc0rqdoqPieBRy/qq+zXGBbyONqcJn2d7PAZW8k5XsWYcSLY/+m8ajAy5+J22yp8VsflpKBfedLQPRNF9FTeZbCJWzcL4333cbZCKnK1xSdspKoRyQXSuz/pb9hBzJl0IpFhQCgZ9E4z7lFcIgZdI2NNlO3EtYu0O8HXrilbHyl9Eo1C1Vo1TK0K52u5FQlPm/IrdtfXcic/3k5kHqwqxTRie2fdFf++6DLLCiZ5iMpuN+7grkY4Upm/KTmAUospvW7imbRLSslNah38RTBFXyDLRfO6KCwhcV+aPu20D7A/UreSuV2uuCdvavHSnUXkWmSLGdSQrY62LkiT2Gys3p/CWCpfyFWWmSwFtNFsovpiDEOdF28f3l7MYXoFCQyGEU8RYLG/zs0Lc7IylD/SMrKZlZ6BzkXzALVlam7PfHJE4W0yQexOd8NOD8cjL4NBpPBsvL5ZIObnGxOFn+IVt56ox+i06+Q0tfuhsob7d0tZvOYudeUkhZ3b9XFvaOmP8FUEsHCO246HhOAgAANAUAAFBLAwQUAAgICADQALJcAAAAAAAAAAAAAAAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWzdlU2P2jAQhu/9FZbvXROyIECEFQWiHlbqgbb3wXESL7YT2d7d8u9rnAD5qraqKlXbXPCMn3k945mQ5cMPKdAL04YXKsLB3QgjpmiRcJVF+NvX+OMMI2NBJSAKxSJ8YgY/rD4sYWFzJhly4cosIMK5teWCEEOdG8xdUTLl9tJCS7DO1BlJNLw6WSnIeDSaEglc4Tpe/058kaacsm1BnyVTthLRTIB1qZuclwYjBdLl+MWDeHVJcifYOcKcHVToPfWZV+wjP2jWCEiOwfnH6OywERq9gIjwyD+YrJbkCgjb52L/1FwNJMfxW3rjSq/PdfQ8AJS6UvpnB7N1OAprtgFVy4Ecwul83eYb+mGPX4fhrqMf3vj7Hj9zdEf//sZPevxmPt9c76QBVcvpAD8Ogl2L91AuuDoO3vjuQl+RtBCfB/HJJFjPPtX4jSKN8anilW0NU2OOJDwVOnaAb66bUYXsqWQpUMetNQeBUcktzWOQXJxcihjRHLRh1jXzfDQsGDRituwJvj+jPSjzdiQ1fxZJOolLrt5pFbfESbNRvm2yaXAh9vYk2KPxRZpC8CR2Tm947DoWZe6W2CtedyqrFfTPFUi/LKHaFnqN8DScnK8OyginrrduKcskwkZlGIHI3OeAWu2HudTGbsHkVQr+pKpDklum6/8n9T6VSfdyWJoyan/huZlurxIZ3P37MBnK7JDF/+f8dgsjrdeW9D7sF8/qJ1BLBwjk/1WAIQIAANEIAABQSwMEFAAICAgA0ACyXAAAAAAAAAAAAAAAAA0AAAB4bC9zdHlsZXMueG1s7VhPT9swHL3vU1i+jyQlFJjSIMbUaZcJjSIhTTuYxEks/CeyXWj49Ps5TtOEwiZ1hxWpJ9svv/f88uyodpOLleDokWrDlJzh6CjEiMpM5UyWM3y7mH88w8hYInPClaQz3FCDL9IPibENpzcVpRaBgjQzXFlbfwoCk1VUEHOkairhSaG0IBaGugxMrSnJjSMJHkzCcBoIwiROE7kUc2ENytRSWrDRQ8g333IApzFGXu5K5WDlK5VUE46DNAk6gTQplNzoxNgDaWKe0SPhIBK6ckkE9eNLzbxCQQTjjQcnraQn7kAP94beNi4UxnkfygR7IE1qYi3Vcg4D1PUXTQ3JSlhqL9PW/aW61KSJJicDQtvAvPdK57C1hsvqIZQzUipJ+G09wwXhhuIe+qKe5BpME04LC8KalZVrraoDJ2KtEtBZc9zUXrnvwPQZ5fzG7dO7YvP2IYiuiu19JdsBbH/nvet6pW5A6po3c+VErF7SDvjcloygS85KKeiLwmutLM1s+5m1cJqQdSGqlGbPIO0WsOy2tfsqLcsc5N8XI0tX9oeyxKuApydN6gWAfYhM5u3E8MxUmsmHhZqz/jHEVPc2EFfZA83XJiuWA3VQGayKF0mFm5yiXXPqfL4MaggPk1pvg/djZnIw84aZnb+tg5mDmYOZg5mDmV3MxMf79EsZR3vlJt4rN5N9cnP+n80Ew+O7P8wPzvHRrsf4VbHtfOjnH62/gzN90EU5uCD1sU7xAEXuqjnD392dmw+Su18ybpn0o2CbcKWEIOv66GREOH6TgH6Gv3rSdESavkpaak1l1vSc0xEn/hNnNNfZiHf6Gu+a6gzWoKecjyj+6rsJEwabv0fS31BLBwiMT4YUgwIAAGMRAABQSwMEFAAICAgA0ACyXAAAAAAAAAAAAAAAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWydVU1z2zgMve+v0OjQ09ay3TppWtudjLPediaNM3G6ndkbLUIWJyTBkpSd5NcvSH3W2UOmPtgiID4A7wHw/POjkskBrBOoF+lkNE4T0DlyofeL9Pv9+u2HNHGeac4kalikT+DSz8s/5ke0D64E8AkBaLdIS+/NxyxzeQmKuREa0OQp0Crm6Wj3mTMWGI+XlMym4/FZppjQaY3w0b4GA4tC5HCFeaVA+xrEgmSe0nelMK5Fe+SvwuOWHanUNp9Bile1p8ObvH+Bp0Ru0WHhRzmqJrWXVV5kF7/U+Winv4c0mVGpBxGUmrZgKn9NlYrZh8q8JWxDTO2EFP4pFpwu5xH/1iaFkB7sN+QkcsGkA/IZtoct+O8m+v093pKhdWfLedZcXs65ID1CZomFYpFeToI7ev8RcHSD58SVeFxTcpVkrsWKxr+t4NdCA1m9rRrjHR5XKL8QEdSjQ8e/QIy1Biv2JaV3DYXvID3bbUFC7oEP720qLynI9kntUHYAHApWSR9SoHBoW/uBMl6kOnApCRJNCLECKWONSR7e/Ur4Z+/T5BlRbXMmiaHJeDw438Trp9bA5TV7wirS0njDWO0QH4Ip4I6DQrGKwK1hYQSbLNKEkfUAfTb9ub6auJ8DNbJOguFzK806tgvp3DBBLPwQ3JeU12Q0ezebzM6ms44nUuULBM7JPR3Risgr51G1tkaBZ5LoxIQ1+9dwAEkgMcuhjcLWRWe/ZNUkecU8o0qMFdpvTJz5pKTWoDntW2nft9GphXq5FbZEK55ReyZXtEjAhiZpXqdt6EX+0pHVA/GN2b2gwDI223h0/uF81nRgfySN4jadTc+7D9G0Q080/Z+njB3eAxSIfnDOumGsDLWBAbsVz9QLFyTyoOXikLa6NcdOqDQJEBsb43A86vsS9IaqJQmsoGLjFl2kBq23TFCD7STLHy41/1EK3819QjtzMGY5tdsKVVjHLkyKhhDXOh/6+6ZSuxCNYlcO1qfmUymujFik70IhrQa9JUcjgqaxl2u21pGjhIuiIJ20j/h9mq15w/lfB9D9WkPO64WyfMOU+bSK329+Vug/3dMec8kNrak7VEz/eQd72lO2dsb3JtP4cznPepiAWCfze4iBkyQ+30bYBmueDeukY/enu/wPUEsHCP7XgjRbAwAAuAcAAFBLAwQUAAgICADQALJcAAAAAAAAAAAAAAAACwAAAF9yZWxzLy5yZWxzrZLPSgMxEIfvfYqQe3e2FURks72I0JtIfYCYzP5hN5kwGXV9e4MIWqmlB49JfvPNN0Oa3RJm9YqcR4pGb6paK4yO/Bh7o58O9+sbvWtXzSPOVkokD2PKqtTEbPQgkm4Bshsw2FxRwlheOuJgpRy5h2TdZHuEbV1fA/9k6PaIqfbeaN77jVaH94SXsKnrRod35F4CRjnR4leikC33KEYvM7wRT89EU1WgGk67bC93+XtOCCjWW7HgiHGduFSzjJi/dTy5h3KdPxPnhK7+czm4CEaP/rySTenLaNXA0SdoPwBQSwcIZqqCt+AAAAA7AgAAUEsDBBQACAgIANAAslwAAAAAAAAAAAAAAAARAAAAZG9jUHJvcHMvY29yZS54bWxtUsluwjAQvfcrIt8TO1AQjZIgtRWnIlUC1Ko31x6C28Sx7IHA39dJSujCbd7iN2OP0/mxKoMDWKdqnZE4YiQALWqpdJGRzXoRzkjgkGvJy1pDRk7gyDy/SYVJRG3h2dYGLCpwgQ/SLhEmIztEk1DqxA4q7iLv0F7c1rbi6KEtqOHikxdAR4xNaQXIJUdO28DQDInkO1KKIdLsbdkFSEGhhAo0OhpHMb14EWzlrh7olB/OSuHJwFXrWRzcR6cGY9M0UTPurH7+mL4un1bdVUOl26cSQPL0e5BEWOAIMvABSd/urLyMHx7XC5KP2GgaskkYz9bsLmGTZHL7ltI/59vAvq5t3qoX4GsJTlhl0O+wF38RHpdcF3v/4DnocLPqLAPVrrLkDpd+6VsF8v7kM65wnrJwUO1HyVnnGGDbwu3fP0Bg338AvkaFJfT0ufz3efIvUEsHCC4C7k5RAQAAiAIAAFBLAwQUAAgICADQALJcAAAAAAAAAAAAAAAAEAAAAGRvY1Byb3BzL2FwcC54bWydkbtuwjAUhvc+RWR1JXZMnKbIMapUdUJqhxR1Q8Y+AVeJbcUGwdvXgArMnOnc9P3nwueHoc/2MAbjbIOKnKAMrHLa2E2DvtuPSY2yEKXVsncWGnSEgObiiX+NzsMYDYQsEWxo0DZGP8M4qC0MMuSpbFOlc+MgYwrHDXZdZxS8O7UbwEZMCakwHCJYDXrir0B0Ic728VGoduo0X1i2R594grcw+F5GEBzf3NZF2bdmAEFS+hrwN+97o2RMFxELsx7h8yyBaZXTfJrT54Wxu8Pqp65WVZnddazSDr+gItalZqDLl3VF6rKgrKO1ptWrZBUjBWGalVoDmZYc32udhJeXT4iC5STZueE/x/Ht6OIPUEsHCHwJxBAPAQAAuQEAAFBLAwQUAAgICADQALJcAAAAAAAAAAAAAAAAEwAAAFtDb250ZW50X1R5cGVzXS54bWytVMlugzAQvecrkK8ROOmhqipIDl2ObQ7pB7h4ADd4ke2k5O87NmmkppQoSi9YeN42g3G+7GSb7MA6oVVB5tmMJKBKzYWqC/K2fk7vyHIxydd7Ay5BrHIFabw395S6sgHJXKYNKKxU2krm8dXW1LByw2qgN7PZLS218qB86oMGWeSPULFt65OnDrd730p5zjwjyUOPDXYFYca0omQeIbRLK30A0UGJDwP1CV/IkCEWhjlGDVPC/jDDQutGUu4UPxlHehhFhsyIcY0wboqAPxxCZWQMPe8Vv5gVHJIVs/6FSUTRrqWf2m7etd5k4yIDKXVViRK4LrcSKZkzFhh3DYCXbRbXTDKhpuP+Hg8E9M/51RmizBlD5/ctuP9uN4qecQ6jjgRH43J9vz9DHPXHciB3ZbVx+ItZuDzA99EM7NSgEFgvxjs/OqL01R1DOPUc+G/vSU7jjbP4AlBLBwj+Xk+tWAEAAKAEAABQSwECFAAUAAgICADQALJc+83WB80AAAAcAgAAGgAAAAAAAAAAAAAAAAAAAAAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAAUAAgICADQALJc7bjoeE4CAAA0BQAADwAAAAAAAAAAAAAAAAAVAQAAeGwvd29ya2Jvb2sueG1sUEsBAhQAFAAICAgA0ACyXOT/VYAhAgAA0QgAABMAAAAAAAAAAAAAAAAAoAMAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAAUAAgICADQALJcjE+GFIMCAABjEQAADQAAAAAAAAAAAAAAAAACBgAAeGwvc3R5bGVzLnhtbFBLAQIUABQACAgIANAAslz+14I0WwMAALgHAAAYAAAAAAAAAAAAAAAAAMAIAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAAUAAgICADQALJcZqqCt+AAAAA7AgAACwAAAAAAAAAAAAAAAABhDAAAX3JlbHMvLnJlbHNQSwECFAAUAAgICADQALJcLgLuTlEBAACIAgAAEQAAAAAAAAAAAAAAAAB6DQAAZG9jUHJvcHMvY29yZS54bWxQSwECFAAUAAgICADQALJcfAnEEA8BAAC5AQAAEAAAAAAAAAAAAAAAAAAKDwAAZG9jUHJvcHMvYXBwLnhtbFBLAQIUABQACAgIANAAslz+Xk+tWAEAAKAEAAATAAAAAAAAAAAAAAAAAFcQAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAJAAkAPgIAAPARAAAAAA==")),
				Boolean.TRUE);
		//
	}

	@Test
	void testIsXls() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertEquals(invoke(METHOD_IS_XLS, null, decode(decoder,
				"0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAOwADAP7/CQAGAAAAAAAAAAAAAAABAAAACAAAAAAAAAAAEAAAAgAAAAEAAAD+////AAAAAAAAAAD////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9//////////7///8EAAAABQAAAAYAAAAHAAAA/v///wkAAAD+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP7///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAAQAAAAIAAAADAAAABAAAAAUAAAAGAAAABwAAAAgAAAAJAAAACgAAAAsAAAAMAAAADQAAAA4AAAAPAAAAEAAAABEAAAASAAAAEwAAABQAAAAVAAAAFgAAABcAAAAYAAAA/v///xoAAAD+/////v///x0AAAAeAAAA/v///yAAAAD+//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8JCBAAAAYFALsNzAcAAAAABgAAAOEAAgCwBMEAAgAAAOIAAABcAHAABAAAQ2FsYyAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIEIAAgCwBGEBAgAAAMABAAA9AQIAAQCcAAIADgCvAQIAAAC8AQIAAAA9ABIAAAAAAABAACA4AAAAAAABAPQBQAACAAAAjQACAAAAIgACAAAADgACAAEAtwECAAAA2gACAAAAMQAaAMgAAAD/f5ABAAAAAgAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAHgQMAKQABwAAR2VuZXJhbOAAFAAAAKQA9f8gAAAAAAAAAAAAAADAIOAAFAABAAAA9f8gAAD0AAAAAAAAAADAIOAAFAABAAAA9f8gAAD0AAAAAAAAAADAIOAAFAACAAAA9f8gAAD0AAAAAAAAAADAIOAAFAACAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAA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")),
				Boolean.TRUE);
		//
	}

	@Test
	void testWrite() throws IOException, IllegalAccessException, InvocationTargetException {
		//
		try (final Writer writer = new OutputStreamWriter(new ProcessBuilder("date").start().getOutputStream())) {
			//
			Assert.assertNull(invoke(METHOD_WRITE, null, writer, null));
			//
		} // try
			//
	}

	@Test
	void testToURI() throws Throwable {
		//
		Assert.assertNotNull(invoke(METHOD_TO_URI, null, toFile(Path.of("."))));
		//
	}

	@Test
	void testGetPage() throws IllegalAccessException, InvocationTargetException {
		//
		final PDDocument pdDocument = new PDDocument();
		//
		Assert.assertNull(invoke(METHOD_GET_PAGE, null, pdDocument, 0));
		//
		final PDPage pdPage = new PDPage();
		//
		pdDocument.addPage(pdPage);
		//
		Assert.assertEquals(pdPage, invoke(METHOD_GET_PAGE, null, pdDocument, 0));
		//
	}

	@Test
	void testTestAndGet() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNull(invoke(METHOD_TEST_AND_GET, null, Boolean.TRUE, null, null));
		//
	}

	@Test
	void testRemoveRow() throws IllegalAccessException, InvocationTargetException {
		//
		final DefaultTableModel dtm = new DefaultTableModel();
		//
		Assert.assertNull(invoke(METHOD_REMOVE_ROW, null, dtm, 0));
		//
		dtm.addRow(new Object[] {});
		//
		Assert.assertNull(invoke(METHOD_REMOVE_ROW, null, dtm, 0));
		//
	}

	@Test
	void testTestAndGetAsBoolean() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertEquals(invoke(METHOD_TEST_AND_GET_AS_BOOLEAN, null, Boolean.TRUE, null), Boolean.FALSE);
		//
		if ((ih = ObjectUtils.getIfNull(ih, IH::new)) != null) {
			//
			ih.getAsBoolean = Boolean.FALSE;
			//
		} // if
			//
		final BooleanSupplier booleanSupplier = Reflection.newProxy(BooleanSupplier.class, ih);
		//
		Assert.assertEquals(invoke(METHOD_TEST_AND_GET_AS_BOOLEAN, null, Boolean.TRUE, booleanSupplier), Boolean.FALSE);
		//
		if (ih != null) {
			//
			ih.getAsBoolean = Boolean.TRUE;
			//
		} // if
			//
		Assert.assertEquals(invoke(METHOD_TEST_AND_GET_AS_BOOLEAN, null, Boolean.TRUE, booleanSupplier), Boolean.TRUE);
		//
	}

	@Test
	void testReplace() throws IllegalAccessException, InvocationTargetException {
		//
		String string = "";
		//
		Assert.assertSame(string, invoke(METHOD_REPLACE, null, string, null, null));
		//
		Assert.assertSame(string, invoke(METHOD_REPLACE, null, string, List.of(string), null));
		//
		Assert.assertEquals(string = "12", invoke(METHOD_REPLACE, null, string, List.of(string), string));
		//
	}

	@Test
	void testAddRow() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNull(invoke(METHOD_ADD_ROW, null, new DefaultTableModel(), null));
		//
	}

	@Test
	void testGetParentFile() throws Throwable {
		//
		Assert.assertNull(invoke(METHOD_GET_PARENT_FILE, null, toFile(Path.of("."))));
		//
	}

	@Test
	void testTestAndAccept() throws IllegalAccessException, InvocationTargetException {
		//
		final Predicate<?> alwaysTrue = Predicates.alwaysTrue();
		//
		Assert.assertNull(invoke(METHOD_TEST_AND_ACCEPT3, null, alwaysTrue, null, null));
		//
		final FailableConsumer<?, ?> failableConsumer = Reflection.newProxy(FailableConsumer.class,
				ih = ObjectUtils.getIfNull(ih, IH::new));
		//
		Assert.assertNull(invoke(METHOD_TEST_AND_ACCEPT3, null, alwaysTrue, null, failableConsumer));
		//
		if (ih != null) {
			//
			ih.throwable = new Throwable();
			//
		} // if
			//
		Assert.assertNull(invoke(METHOD_TEST_AND_ACCEPT4_PREDICATE, null, alwaysTrue, null, failableConsumer, null));
		//
		if (ih != null) {
			//
			ih.test = Boolean.TRUE;
			//
		} // if
			//
		final FailableBiPredicate<?, ?, ?> failableBiConsumer = Reflection.newProxy(FailableBiPredicate.class, ih);
		//
		Assert.assertNull(
				invoke(METHOD_TEST_AND_ACCEPT4_FAILABLE_BI_PREDICATE, null, failableBiConsumer, null, null, null));
		//
		Assert.assertNull(invoke(METHOD_TEST_AND_ACCEPT4_FAILABLE_BI_PREDICATE, null, failableBiConsumer, null, null,
				Reflection.newProxy(FailableBiConsumer.class, ih)));
		//
	}

	@Test
	void testAnd() throws IllegalAccessException, InvocationTargetException {
		//
		final Predicate<?> alwaysTrue = Predicates.alwaysTrue();
		//
		Assert.assertEquals(invoke(METHOD_AND, null, null, alwaysTrue, null), Boolean.FALSE);
		//
		Assert.assertEquals(invoke(METHOD_AND, null, null, alwaysTrue, alwaysTrue), Boolean.TRUE);
		//
	}

	@Test
	void testCreateStringTextPositionEntryMap() throws IllegalAccessException, InvocationTargetException {
		//
		final Sheet sheet = Reflection.newProxy(Sheet.class, ih = ObjectUtils.getIfNull(ih, IH::new));
		//
		Assert.assertNull(invoke(METHOD_CREATE_STRING_TEXT_POSITION_ENTRY_MAP, null, sheet, null));
		//
		final Row row = Reflection.newProxy(Row.class, ih);
		//
		if (ih != null) {
			//
			ih.iterator = iterator(List.of(row));
			//
		} // if
			//
		Assert.assertNull(invoke(METHOD_CREATE_STRING_TEXT_POSITION_ENTRY_MAP, null, sheet, null));
		//
		if (ih != null) {
			//
			ih.iterator = iterator(List.of(row));
			//
			ih.cell = Reflection.newProxy(Cell.class, ih);
			//
			ih.stringCellValue = "";
			//
		} // if
			//
		Assert.assertEquals(
				JsonMapper.builder().changeDefaultVisibility(x -> x != null ? x.withFieldVisibility(Visibility.ANY) : x)
						.build()
						.writeValueAsString(invoke(METHOD_CREATE_STRING_TEXT_POSITION_ENTRY_MAP, null, sheet, null)),
				"{\"\":{\"file\":null,\"marker\":\"\",\"sheetName\":null,\"text\":\"\",\"textPosition\":null}}");
		//
	}

	private static <T> Iterator<T> iterator(final Iterable<T> instance) {
		return instance != null ? instance.iterator() : null;
	}

	@Test
	void testAddPDAnnotations() throws IllegalAccessException, InvocationTargetException, ClassNotFoundException {
		//
		final Map<?, Object> map = new LinkedHashMap<>();
		//
		map.put(null, null);
		//
		Assert.assertEquals(invoke(METHOD_ADD_PD_ANNOTATIONS, null, map, null, null, Boolean.TRUE), Boolean.TRUE);
		//
		final Class<?> clz = Class.forName("org.apache.pdfbox.AddAudioJPanel$TextPositionEntry");
		//
		final Object textPositionEntry = Narcissus.allocateInstance(clz);
		//
		map.put(null, textPositionEntry);
		//
		Assert.assertEquals(invoke(METHOD_ADD_PD_ANNOTATIONS, null, map, null, null, Boolean.TRUE), Boolean.TRUE);
		//
		FieldUtils.writeDeclaredField(textPositionEntry, "textPosition", Narcissus.allocateInstance(TextPosition.class),
				true);
		//
		Assert.assertEquals(invoke(METHOD_ADD_PD_ANNOTATIONS, null, map, null, null, Boolean.TRUE), Boolean.TRUE);
		//
	}

	@Test
	void testToPath() throws Throwable {
		//
		Assert.assertNotNull(invoke(METHOD_TO_PATH, null, toFile(Path.of("."))));
		//
	}

	@Test
	void testGetAbsoluteFile() throws Throwable {
		//
		Assert.assertNotNull(invoke(METHOD_GET_ABSOLUTE_FILE, null, toFile(Path.of("."))));
		//
	}

	@Test
	void testSave() throws Throwable {
		//
		Assert.assertNull(invoke(METHOD_SAVE, null, null, null));
		//
		final PDDocument pdDocument = new PDDocument();
		//
		Assert.assertNull(invoke(METHOD_SAVE, null, pdDocument, null));
		//
		Assert.assertNull(invoke(METHOD_SAVE, null, pdDocument, toFile(Path.of("."))));
		//
		final File file = File.createTempFile(" ".repeat(3), null);
		//
		if (file != null) {
			//
			file.deleteOnExit();
			//
		} // if
			//
		Assert.assertNull(invoke(METHOD_SAVE, null, pdDocument, file));
		//
		if (file != null) {
			//
			file.delete();
			//
		} // if
			//
	}

	@Test
	void testToRuntmeException() throws IllegalAccessException, InvocationTargetException {
		//
		final RuntimeException runtimeException = new RuntimeException();
		//
		Assert.assertSame(runtimeException, invoke(METHOD_TO_RUNTIME_EXCEPTION, null, runtimeException));
		//
	}

	@Test
	void testOpen() throws Throwable {
		//
		Assert.assertNull(invoke(METHOD_OPEN, null, Narcissus.allocateInstance(Desktop.class), toFile(Path.of("."))));
		//
	}

	@Test
	void testIsDirectory() throws Throwable {
		//
		Assert.assertEquals(invoke(METHOD_IS_DIRECTORY, null, toFile(Path.of("pom.xml"))), Boolean.FALSE);
		//
	}

	@Test
	void testSetSubtype() throws IllegalAccessException, InvocationTargetException, IOException {
		//
		try (final InputStream is = Files.newInputStream(Path.of("pom.xml"))) {
			//
			final PDEmbeddedFile pdEmbeddedFile = new PDEmbeddedFile(new PDDocument(), is);
			//
			Assert.assertNull(invoke(METHOD_SET_SUB_TYPE, null, pdEmbeddedFile, null));
			//
		} // try
			//
	}

	@Test
	void testGetMediaBox() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNotNull(invoke(METHOD_GET_MEDIA_BOX, null, new PDPage()));
		//
	}

	@Test
	void testGetHeight() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNotNull(invoke(METHOD_GET_HEIGHT, null, new PDRectangle()));
		//
	}

	@Test
	void testClear() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNotNull(invoke(METHOD_CLEAR, null, Narcissus.allocateInstance(TextStringBuilder.class)));
		//
	}

	@Test
	void testAppend() throws IllegalAccessException, InvocationTargetException {
		//
		final Object object = Narcissus.allocateInstance(TextStringBuilder.class);
		//
		Assert.assertNotNull(invoke(METHOD_APPEND_CHAR, null, object, Character.valueOf(' ')));
		//
		Assert.assertNotNull(invoke(METHOD_APPEND_INT, null, object, Integer.valueOf(0)));
		//
		Assert.assertNotNull(invoke(METHOD_APPEND_STRING, null, object, null));
		//
		Assert.assertNotNull(invoke(METHOD_APPEND_STRING, null, object, Narcissus.allocateInstance(String.class)));
		//
	}

	@Test
	void testGetFileExtension() throws IllegalAccessException, InvocationTargetException {
		//
		final ContentInfoUtil ciu = new ContentInfoUtil();
		//
		Assert.assertEquals(
				invoke(METHOD_GET_FILE_EXTENSION, null,
						ciu.findMatch(decode(decoder, "UklGRiQAAABXQVZFZm10IBAAAAABAAEARKwAAIhYAQACABAAZGF0YQAAAAA="))),
				"wav");
		//
		Assert.assertEquals(invoke(METHOD_GET_FILE_EXTENSION, null, ciu.findMatch(decode(decoder,
				"/+MYxAAAAANIAAAAAExBTUUzLjk4LjIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"))),
				"mp3");
		//
	}

	private static byte[] decode(final Decoder instance, final String string) {
		return instance != null ? instance.decode(string) : null;
	}

	@Test
	void testIif() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNull(invoke(METHOD_IIF, null, Boolean.TRUE, null, null));
		//
	}

	@Test
	void testContainsKey() throws IllegalAccessException, InvocationTargetException {
		//
		final String string = "";
		//
		Assert.assertEquals(invoke(METHOD_CONTAINS_KEY, null, Map.of(string, string), string), Boolean.TRUE);
		//
	}

	@Test
	void testSetModDate() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertNull(invoke(METHOD_SET_MOD_DATE, null, new PDEmbeddedFile(new PDDocument()), null));
		//
	}

	@Test
	void testSetSize() throws IllegalAccessException, InvocationTargetException {
		//
		final PDEmbeddedFile pdEmbeddedFile = new PDEmbeddedFile(new PDDocument());
		//
		Assert.assertNull(invoke(METHOD_SET_SIZE, null, pdEmbeddedFile, Integer.MIN_VALUE));
		//
		Assert.assertNull(invoke(METHOD_SET_SIZE, null, pdEmbeddedFile, Integer.MAX_VALUE));
		//
	}

	@Test
	void testCreateInputStreamWorkbookFailableFunction() throws IllegalAccessException, InvocationTargetException {
		//
		// xlsx
		//
		Assert.assertNotNull(invoke(METHOD_CREATE_INPUT_STREAM_WORK_BOOK_FAILABLE_FUNCTION, null, decode(decoder,
				"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")));
		//
		// xls
		//
		Assert.assertNotNull(invoke(METHOD_CREATE_INPUT_STREAM_WORK_BOOK_FAILABLE_FUNCTION, null, decode(decoder,
				"0M8R4KGxGuEAAAAAAAAAAAAAAAAAAAAAOwADAP7/CQAGAAAAAAAAAAAAAAABAAAACAAAAAAAAAAAEAAAAgAAAAEAAAD+////AAAAAAAAAAD////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9//////////7///8EAAAABQAAAAYAAAAHAAAA/v///wkAAAD+/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////1IAbwBvAHQAIABFAG4AdAByAHkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWAAUA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD+////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///////////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP7///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////////////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/v///wAAAAAAAAAAAQAAAAIAAAADAAAABAAAAAUAAAAGAAAABwAAAAgAAAAJAAAACgAAAAsAAAAMAAAADQAAAA4AAAAPAAAAEAAAABEAAAASAAAAEwAAABQAAAAVAAAAFgAAABcAAAAYAAAA/v///xoAAAD+/////v///x0AAAAeAAAA/v///yAAAAD+//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8JCBAAAAYFALsNzAcAAAAABgAAAOEAAgCwBMEAAgAAAOIAAABcAHAABAAAQ2FsYyAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIEIAAgCwBGEBAgAAAMABAAA9AQIAAQCcAAIADgCvAQIAAAC8AQIAAAA9ABIAAAAAAABAACA4AAAAAAABAPQBQAACAAAAjQACAAAAIgACAAAADgACAAEAtwECAAAA2gACAAAAMQAaAMgAAAD/f5ABAAAAAgAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAMQAaAMgAAAD/f5ABAAAAAAAABQFBAHIAaQBhAGwAHgQMAKQABwAAR2VuZXJhbOAAFAAAAKQA9f8gAAAAAAAAAAAAAADAIOAAFAABAAAA9f8gAAD0AAAAAAAAAADAIOAAFAABAAAA9f8gAAD0AAAAAAAAAADAIOAAFAACAAAA9f8gAAD0AAAAAAAAAADAIOAAFAACAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAAFAAAAAAA9f8gAAD0AAAAAAAAAADAIOAA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")));
		//
	}

	@Test
	void testIsSelected() throws IllegalAccessException, InvocationTargetException {
		//
		final AbstractButton ab = new JButton();
		//
		ab.setSelected(true);
		//
		Assert.assertEquals(invoke(METHOD_IS_SELECTED, null, ab), Boolean.TRUE);
		//
	}

	@Test
	void testOr() throws IllegalAccessException, InvocationTargetException {
		//
		FailablePredicate<?, ?> failablePredicate1 = Objects::nonNull;
		//
		Assert.assertEquals(invoke(METHOD_OR, null, "", failablePredicate1, null), Boolean.TRUE);
		//
		if (ih != null) {
			//
			ih.test = Boolean.TRUE;
			//
		} // if
			//
		Assert.assertEquals(
				invoke(METHOD_OR, null, null, failablePredicate1, Reflection.newProxy(FailablePredicate.class, ih)),
				Boolean.TRUE);
		//
	}

	@Test
	void testIsValidTypstFile() throws IllegalAccessException, InvocationTargetException {
		//
		Assert.assertEquals(invoke(METHOD_IS_VALID_TYPST_FILE, null, new File("pom.xml")), Boolean.FALSE);
		//
	}

	@Test
	void testGetPages() throws IllegalAccessException, InvocationTargetException, IOException {
		//
		try (final PDDocument pdDocument = new PDDocument()) {
			//
			Assert.assertNotNull(invoke(METHOD_GET_PAGES, null, pdDocument));
			//
		} // try
			//
	}

	@Test
	void testIndexOf() throws IllegalAccessException, InvocationTargetException, IOException {
		//
		try (final PDDocument pdDocument = new PDDocument()) {
			//
			final PDPageTree pdPageTree = pdDocument.getPages();
			//
			Assert.assertEquals(invoke(METHOD_INDEX_OF, null, pdPageTree, null), Integer.valueOf(-1));
			//
			Assert.assertEquals(invoke(METHOD_INDEX_OF, null, pdPageTree, new PDPage()), Integer.valueOf(-1));
			//
		} // try
			//
	}

	@Test
	void testCreateTypstFileFilter() throws Throwable {
		//
		final FileFilter fileFilter = cast(FileFilter.class, invoke(METHOD_CREATE_TYPST_FILE_FILTER, null));
		//
		if (fileFilter == null) {
			//
			return;
			//
		} // if
			//
		Assert.assertTrue(fileFilter.accept(toFile(Path.of("."))));
		//
		Assert.assertEquals(fileFilter.accept(toFile(Path.of("sample", "sample.typ"))), exists("TYPST"));
		//
	}

	private static boolean exists(final String command) throws Throwable {
		try {
			final Object obj = invoke(METHOD_EXISTS_STRING, null, command);
			if (obj instanceof Boolean) {
				return ((Boolean) obj).booleanValue();
			} // if
			throw new Throwable(Objects.toString(getClass(obj)));
		} catch (final InvocationTargetException e) {
			throw e.getTargetException();
		}
	}

	@Test
	void testActionPerformed() throws Throwable {
		//
		if (instance == null) {
			//
			return;
			//
		} // if
			//
			// btnFileSpreadsheet
			//
		final AbstractButton btnFileSpreadsheet = new JButton();
		//
		FieldUtils.writeDeclaredField(instance, "btnFileSpreadsheet", btnFileSpreadsheet, true);
		//
		instance.actionPerformed(new ActionEvent(btnFileSpreadsheet, 0, null));
		//
		// btnExecute
		//
		final AbstractButton btnExecute = new JButton();
		//
		FieldUtils.writeDeclaredField(instance, "btnExecute", btnExecute, true);
		//
		ActionEvent actionEvent = new ActionEvent(btnExecute, 0, null);
		//
		instance.actionPerformed(actionEvent);
		//
		// btnBrowse
		//
		final AbstractButton btnBrowse = new JButton();
		//
		FieldUtils.writeDeclaredField(instance, "btnBrowse", btnBrowse, true);
		//
		instance.actionPerformed(actionEvent);
		//
		instance.actionPerformed(actionEvent = new ActionEvent(btnBrowse, 0, null));
		//
		final JTextComponent tfFilePdf = new JTextField();
		//
		tfFilePdf.setText(Objects.toString(toFile(Path.of(".")).getAbsolutePath()));
		//
		FieldUtils.writeDeclaredField(instance, "tfFilePdf", tfFilePdf, true);
		//
		instance.actionPerformed(actionEvent);
		//
	}

	@Test
	void testGetTextLocation() throws Throwable {
		//
		final Class<?> clz = Class.forName("org.apache.pdfbox.AddAudioJPanel$GetTextLocation");
		//
		final Object object = Narcissus.allocateInstance(clz);
		//
		final Method writeString = clz != null ? clz.getDeclaredMethod("writeString", String.class, List.class) : null;
		//
		Assert.assertNull(Narcissus.invokeMethod(object, writeString, null, null));
		//
		final Collection<?> list = new ArrayList<>();
		//
		list.add(null);
		//
		Assert.assertNull(Narcissus.invokeMethod(object, writeString, null, list));
		//
		final Object textPosition = Narcissus.allocateInstance(TextPosition.class);
		//
		Assert.assertNull(Narcissus.invokeMethod(object, writeString, null, List.of(textPosition)));
		//
		final Map<?, ?> map = cast(Map.class, FieldUtils.readDeclaredField(object, "map", true));
		//
		if (map != null) {
			//
			map.put(null, null);
			//
		} // if
			//
		Assert.assertNull(Narcissus.invokeMethod(object, writeString, null, List.of(textPosition)));
		//
		Assert.assertNull(Narcissus.invokeStaticMethod(
				clz != null ? clz.getDeclaredMethod("get", Map.class, Object.class) : null, null, null));
		//
	}

	@Test
	void testTextPositionEntry() throws ClassNotFoundException, NoSuchMethodException {
		//
		final Class<?> clz = Class.forName("org.apache.pdfbox.AddAudioJPanel$TextPositionEntry");
		//
		Assert.assertNull(Narcissus.invokeStaticMethod(
				clz != null ? clz.getDeclaredMethod("setTextPosition", clz, TextPosition.class) : null, null, null));
		//
	}

}