package rebound.richsheets.impls.live.googlesheets;

import java.io.CharArrayWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.Charset;
import java.nio.charset.CharsetDecoder;
import java.nio.charset.CodingErrorAction;
import javax.annotation.Nonnull;

public class ReboundStaticallyCopiedUtilities
{
	@Nonnull
	public static String readAllText(File file) throws IOException
	{
		try (InputStream in = new FileInputStream(file))
		{
			return readAllText(in, Charset.defaultCharset());
		}
	}
	
	public static String readAllText(InputStream in, Charset encoding) throws IOException
	{
		if (encoding == null)
			encoding = Charset.defaultCharset();
		
		CharsetDecoder decoder = encoding.newDecoder();
		decoder.onUnmappableCharacter(CodingErrorAction.REPORT);
		decoder.onMalformedInput(CodingErrorAction.REPORT);
		
		return new String(readAll(new InputStreamReader(in, decoder)));
	}
	
	public static char[] readAll(Reader in) throws IOException
	{
		CharArrayWriter buff = new CharArrayWriter();
		pump(in, buff);
		return buff.toCharArray();
	}
	
	public static long pump(Reader in, Writer out) throws IOException
	{
		return pump(in, out, 4096);
	}
	
	public static long pump(Reader in, Writer out, int bufferSize) throws IOException
	{
		long total = 0;
		char[] buffer = new char[bufferSize];
		int amt = in.read(buffer);
		while (amt >= 0)
		{
			total += amt;
			out.write(buffer, 0, amt);
			amt = in.read(buffer);
		}
		return total;
	}
}
