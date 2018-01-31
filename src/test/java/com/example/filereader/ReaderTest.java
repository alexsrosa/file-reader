package com.example.filereader;

import com.example.filereader.util.Reader;
import org.apache.commons.io.FileUtils;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import static org.hamcrest.Matchers.containsString;

public class ReaderTest {

    @Test
    public void givenFileNameAsAbsolutePath_whenUsingClasspath_thenFileData() throws IOException {
        String expectedData = "Hello World from fileTest.txt!!!";

        Class clazz = ReaderTest.class;
        InputStream inputStream = clazz.getResourceAsStream("/fileTest.txt");
        String data = Reader.readFromInputStream(inputStream);

        Assert.assertThat(data, containsString(expectedData));
    }

    @Test
    public void givenFilePath_whenUsingFilesLines_thenFileData() throws IOException, URISyntaxException {
        String expectedData = "Hello World from fileTest.txt!!!";

        Path path = Paths.get(getClass().getClassLoader()
                .getResource("fileTest.txt").toURI());

        StringBuilder data = new StringBuilder();
        Stream<String> lines = Files.lines(path);
        lines.forEach(line -> data.append(line).append("\n"));
        lines.close();

        Assert.assertEquals(expectedData, data.toString().trim());
    }

    @Test
    public void givenFileName_whenUsingFileUtils_thenFileData() throws IOException {
        String expectedData = "Hello World from fileTest.txt!!!";

        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File(classLoader.getResource("fileTest.txt").getFile());
        String data = FileUtils.readFileToString(file);

        Assert.assertEquals(expectedData, data.trim());
    }
}
