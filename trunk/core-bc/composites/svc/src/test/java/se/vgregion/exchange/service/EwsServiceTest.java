package se.vgregion.exchange.service;

import com.microsoft.schemas.exchange.services._2006.messages.*;
import com.microsoft.schemas.exchange.services._2006.types.CalendarItemType;
import com.microsoft.schemas.exchange.services._2006.types.MessageType;
import org.apache.geronimo.mail.util.Base64;
import org.eclipse.jetty.server.Request;
import org.eclipse.jetty.server.Server;
import org.eclipse.jetty.server.handler.AbstractHandler;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.springframework.test.util.ReflectionTestUtils;
import se.vgregion.ldapservice.LdapService;
import se.vgregion.ldapservice.LdapUser;
import se.vgregion.ldapservice.SimpleLdapUser;

import javax.servlet.ServletException;
import javax.servlet.ServletInputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.xml.namespace.QName;
import javax.xml.soap.*;
import java.io.*;
import java.net.URL;
import java.nio.charset.Charset;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.mockito.Matchers.anyString;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

/**
 * @author Patrik Bergström
 */
public class EwsServiceTest {

    private static int port = 8084;

    private EwsService ewsService;
    private ExchangeServicePortType exchangeServicePort;
    private LdapService ldapService;

    @BeforeClass
    public static void beforeClass() throws Exception {
        setupServer();
    }

    @AfterClass
    public static void afterClass() throws Exception {
        server.stop();
        System.out.println("Stopped server...");
    }

    private static Server server = new Server(port);

    public static void setupServer() throws Exception {

        // Setup the server to reply with "{username} {password}" as the client sends the corresponding basic auth
        // credentials.
        AbstractHandler handler = new AbstractHandler() {
            @Override
            public void handle(String s, Request baseRequest, HttpServletRequest request, HttpServletResponse response)
                    throws IOException, ServletException {

                try {
                    String authHeader = baseRequest.getHeader("Authorization");

                    if (authHeader != null && authHeader.startsWith("Basic ")) {
                        String[] up = parseBasic(authHeader.substring(authHeader.indexOf(" ") + 1));
                        String username = up[0];
                        String password = up[1];

                        System.out.println(username + " " + password);

                        if (baseRequest.getMethod().equals("POST")) {
                            String soapAction = request.getHeader("SOAPAction");
                            if (soapAction.contains(
                                    "http://schemas.microsoft.com/exchange/services/2006/messages/FindFolder")) {

                                replyWithSoapDocument(response, "soap-documents/FindFolder-response");
                                return;

                            } else if (soapAction.contains(
                                    "http://schemas.microsoft.com/exchange/services/2006/messages/FindItem")) {

                                ServletInputStream inputStream = request.getInputStream();
                                String soapRequest = streamToString(inputStream);

                                MessageFactory factory = MessageFactory.newInstance();
                                SOAPMessage message = factory.createMessage(new MimeHeaders(),
                                        new ByteArrayInputStream(soapRequest.getBytes(Charset.forName("UTF-8"))));

                                Iterator findItemIterator = message.getSOAPBody().getChildElements(new QName(
                                        "http://schemas.microsoft.com/exchange/services/2006/messages", "FindItem"));

                                SOAPBodyElement findItemElement = (SOAPBodyElement) findItemIterator.next();

                                Iterator calendarViewIterator = findItemElement.getChildElements(
                                        new QName("http://schemas.microsoft.com/exchange/services/2006/messages",
                                                "CalendarView"));

                                if (calendarViewIterator.hasNext()) {
                                    replyWithSoapDocument(response, "soap-documents/CalendarItems-response");
                                } else {
                                    // We assume emails are requested
                                    replyWithSoapDocument(response, "soap-documents/FindUnreadEmails-response");
                                }

                                return;
                            }

                            ServletInputStream inputStream = request.getInputStream();

                            String body = streamToString(inputStream);
                            System.out.println(body);
                        } else {
                            String string = username + " " + password;
                            printToResponse(response, string);
                        }

                        return;
                    }

                    response.setHeader("WWW-Authenticate", "BASIC realm=\"theRealm\"");
                    response.sendError(HttpServletResponse.SC_UNAUTHORIZED, "Please provide username and"
                            + " password");
                } catch (SOAPException e) {
                    e.printStackTrace();
                }

            }
        };

        server.setHandler(handler);

        server.start();
    }

    private static void printToResponse(HttpServletResponse httpServletResponse, String string) throws IOException {
        PrintWriter writer = httpServletResponse.getWriter();
        writer.write(string);
        writer.close();
    }

    private static void replyWithSoapDocument(HttpServletResponse httpServletResponse, String classpathResource)
            throws IOException {
        URL resource = EwsServiceTest.class.getClassLoader().getResource(classpathResource);
        String soapReply = streamToString(resource.openStream());
        printToResponse(httpServletResponse, soapReply);
    }

    @Before
    public void setup() throws IOException {
        ldapService = mock(LdapService.class);

        exchangeServicePort = new ExchangeService(this.getClass().getResource("/wsdl/test-services.wsdl"))
                .getExchangeService();

        ewsService = new EwsService(ldapService, exchangeServicePort);

        ReflectionTestUtils.setField(ewsService, "ewsUser", "theUser");
        ReflectionTestUtils.setField(ewsService, "ewsPassword", "thePassword");
    }

    @Test
    public void testInit() throws Exception {
        ewsService.init();

        // Here we check that the default authenticator is working, i.e. it will authenticate with basic auth.
        URL url = new URL("http://localhost:" + port);

        InputStream inputStream = url.openStream();
        String string = streamToString(inputStream);
        System.out.println(string);

        assertEquals("theUser thePassword", string);
    }

    private static String streamToString(InputStream inputStream) throws IOException {
        BufferedInputStream bis = new BufferedInputStream(inputStream);

        StringBuilder sb = new StringBuilder();
        int n;
        byte[] buf = new byte[1024];
        while (((n = bis.read(buf))) != -1) {
            sb.append(new String(buf, 0, n));
        }

        return sb.toString();
    }

    @Test
    public void testFetchUnreadEmails() throws Exception {
        ewsService.init();

        List<MessageType> emailList = ewsService.fetchUnreadEmails("asdf", 234);

        assertEquals(2, emailList.size()); // The FindUnreadEmails-response file contains two emails.
    }

    @Test
    public void testFetchCalendarEvents() throws Exception {
        ewsService.init();

        List<CalendarItemType> asdf = ewsService.fetchCalendarEvents("asdf", new Date(), new Date());

        assertEquals(2, asdf.size());

        for (CalendarItemType calendarItemType : asdf) {
            assertTrue(calendarItemType instanceof CalendarItemType);
        }
    }

    @Test
    public void testFetchInboxUnreadCount() throws Exception {
        ewsService.init();

        Integer unreadCount = ewsService.fetchInboxUnreadCount("asdfasdf");

        assertEquals((Integer) 17, unreadCount); // The FindFolder-response file says 17
    }

    @Test
    public void testFetchUserSid() throws Exception {

        byte[] bytes = new byte[]{1, 5, 0, 0, 0, 0, 0, 5, 21, 0, 0, 0, 68, 90, -57, -6, 51, 13, 91, 48, 27, -55, 57, 58,
                66, 118, 0, 0};

        SimpleLdapUser user = new SimpleLdapUser("asdf");
        user.setAttributeValue("objectSid", bytes);

        when(ldapService.search(anyString(), anyString())).thenReturn(new LdapUser[]{user});

        String sid = ewsService.fetchUserSid("asdf");

        assertEquals("S-1-5-21-4207368772-811273523-976865563-30274", sid);
    }

    private static String[] parseBasic(String enc) throws UnsupportedEncodingException {
        byte[] bytes = Base64.decode(enc.getBytes());
        String s = new String(bytes, "UTF-8");
        int pos = s.indexOf( ":" );
        if( pos >= 0 )
            return new String[] { s.substring( 0, pos ), s.substring( pos + 1 ) };
        else
            return new String[] { s, null };
    }
}
