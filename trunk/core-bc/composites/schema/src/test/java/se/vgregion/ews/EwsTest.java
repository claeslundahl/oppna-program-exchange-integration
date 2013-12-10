package se.vgregion.ews;

import com.microsoft.schemas.exchange.services._2006.messages.*;
import com.microsoft.schemas.exchange.services._2006.types.*;
import com.sun.security.auth.NTDomainPrincipal;
import org.apache.cxf.configuration.security.AuthorizationPolicy;
import org.apache.cxf.endpoint.Client;
import org.apache.cxf.frontend.ClientProxy;
import org.apache.cxf.helpers.IOUtils;
import org.apache.cxf.interceptor.LoggingInInterceptor;
import org.apache.cxf.interceptor.LoggingOutInterceptor;
import org.apache.cxf.transport.Conduit;
import org.apache.cxf.transport.http.AbstractHTTPDestination;
import org.apache.cxf.transport.http.HTTPConduit;
import org.apache.cxf.transports.http.configuration.HTTPClientPolicy;
import org.junit.Ignore;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.security.auth.kerberos.KerberosPrincipal;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBIntrospector;
import javax.xml.bind.Marshaller;
import javax.xml.namespace.QName;
import javax.xml.ws.BindingProvider;
import javax.xml.ws.Holder;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.net.*;
import java.util.List;

/**
 * @author Patrik Bergström
 */
public class EwsTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(EwsTest.class);

    @Test
    @Ignore
    public void testSoap() throws Exception {
        ExchangeService s = new ExchangeService();

        ExchangeServicePortType exchangeService = s.getExchangeService();

        Client clientProxy = ClientProxy.getClient(exchangeService);
        PrintWriter printWriter = new PrintWriter(System.out, true);
        clientProxy.getOutInterceptors().add(new LoggingOutInterceptor(printWriter));
        clientProxy.getInInterceptors().add(new LoggingInInterceptor(printWriter));
        clientProxy.getInFaultInterceptors().add(new LoggingInInterceptor(printWriter));
        HTTPConduit conduit = (HTTPConduit) clientProxy.getConduit();

        HTTPClientPolicy client = new HTTPClientPolicy();
        client.setAllowChunking(false);
        client.setAutoRedirect(true);
//        client.setProxyServer("127.0.0.1");
//        client.setProxyServerPort(8888);
        conduit.setClient(client);
        /*NTDomainPrincipal ntDomainPrincipal = new NTDomainPrincipal("vgregion.se");

        KerberosPrincipal kerberosPrincipal = new KerberosPrincipal("TK.Exch.Imp.Contact");


        AuthorizationPolicy authorization = new AbstractHTTPDestination.PrincipalAuthorizationPolicy(kerberosPrincipal);
        authorization.setUserName("vgregion.se\\TK.Exch.Imp.Contact");
        authorization.setPassword("WI21aG0O");
        authorization.setAuthorizationType("NTLM");

        conduit.setAuthorization(authorization);*/

        /*AuthorizationPolicy authorization = new AuthorizationPolicy();
        authorization.setUserName("vgregion.se\\TK.Exch.Imp.Contact");
        authorization.setPassword("WI21aG0O");
        ((HTTPConduit) clientProxy.getConduit()).setAuthorization(authorization);*/

        /*conduit.getAuthorization().setAuthorizationType("Negotiate");
        conduit.getAuthorization().setAuthorization("Negotiate");
        conduit.getAuthorization().setUserName("asdf");
        conduit.getAuthorization().setPassword("asdfasdf");*/

//        ((BindingProvider)exchangeService).getRequestContext().put(BindingProvider.USERNAME_PROPERTY, "asdf");
//        ((BindingProvider)exchangeService).getRequestContext().put(BindingProvider.PASSWORD_PROPERTY, "asdf");
// Make NTLM work
        Authenticator.setDefault(new Authenticator() {
            public PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication("vgregion.se\\TK.Exch.Imp.Contact", "WI21aG0O".toCharArray());
            }
        });

        Holder< FindFolderResponseType > findFolderResult = new Holder<FindFolderResponseType>();

        FindFolderType findFolderType = new FindFolderType();
        findFolderType.setTraversal(FolderQueryTraversalType.DEEP);

        IndexedPageViewType indexedPageViewType = new IndexedPageViewType();
        indexedPageViewType.setBasePoint(IndexBasePointType.BEGINNING);
        indexedPageViewType.setOffset(0);
        indexedPageViewType.setMaxEntriesReturned(30);

        RestrictionType restriction = new RestrictionType();
        IsEqualToType equalToType = new IsEqualToType();
        FieldURIOrConstantType displayNameField = new FieldURIOrConstantType();
        ConstantValueType c1 = new ConstantValueType();
        c1.setValue("Inkorg");
        displayNameField.setConstant(c1);
        equalToType.setFieldURIOrConstant(displayNameField);
        PathToUnindexedFieldType fieldType = new PathToUnindexedFieldType();
        fieldType.setFieldURI(UnindexedFieldURIType.FOLDER_DISPLAY_NAME);
        equalToType.setPath(new JAXBElement<BasePathToElementType>(
                new QName("http://schemas.microsoft.com/exchange/services/2006/types", "FieldURI"),
                BasePathToElementType.class, fieldType));

        /*IsGreaterThanType searchExpression = new IsGreaterThanType();
        FieldURIOrConstantType fieldURIOrConstant = new FieldURIOrConstantType();
        ConstantValueType constant = new ConstantValueType();
        constant.setValue("0");
        fieldURIOrConstant.setConstant(constant);
        searchExpression.setFieldURIOrConstant(fieldURIOrConstant);

        PathToUnindexedFieldType basePath = new PathToUnindexedFieldType();
        basePath.setFieldURI("folder:TotalCount");

        searchExpression.setPath(new JAXBElement<BasePathToElementType>(
                new QName("http://schemas.microsoft.com/exchange/services/2006/types", "FieldURI"),
                BasePathToElementType.class, basePath));*/

        JAXBContext context = JAXBContext.newInstance(SearchExpressionType.class);
        JAXBIntrospector jaxbIntrospector = context.createJAXBIntrospector();
        QName qName = new QName("http://schemas.microsoft.com/exchange/services/2006/types", "IsEqualTo");
//        QName qName = new QName("http://schemas.microsoft.com/exchange/services/2006/types", "IsGreaterThan");
        restriction.setSearchExpression(new JAXBElement<SearchExpressionType>(qName, SearchExpressionType.class, equalToType));

        findFolderType.setIndexedPageFolderView(indexedPageViewType);
        findFolderType.setRestriction(restriction);

        NonEmptyArrayOfBaseFolderIdsType folderIdsType = new NonEmptyArrayOfBaseFolderIdsType();
        DistinguishedFolderIdType distinguishedFolderIdType = new DistinguishedFolderIdType();
        distinguishedFolderIdType.setId(DistinguishedFolderIdNameType.ROOT);
        folderIdsType.getFolderIdOrDistinguishedFolderId().add(distinguishedFolderIdType);

        findFolderType.setParentFolderIds(folderIdsType);

        FolderResponseShapeType folderShape = new FolderResponseShapeType();
        folderShape.setBaseShape(DefaultShapeNamesType.ALL_PROPERTIES);

        findFolderType.setFolderShape(folderShape);

        ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
        ConnectingSIDType connectingSID = new ConnectingSIDType();
        connectingSID.setSID("S-1-5-21-4207368772-811273523-976865563-30274");
//        connectingSID.setSID("susro3");
        impersonation.setConnectingSID(connectingSID);

        exchangeService.findFolder(findFolderType, impersonation, null, null,
                null, null, findFolderResult, null);

        System.out.println(findFolderResult.value.getResponseMessages()
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage()
                .size());

        ArrayOfResponseMessagesType responseMessages = findFolderResult.value.getResponseMessages();
        JAXBElement<? extends ResponseMessageType> jaxbElement = responseMessages
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage().get(0);
        Integer unreadCount = ((FolderType) (((((FindFolderResponseMessageType) jaxbElement.getValue()).getRootFolder()).getFolders()).getFolderOrCalendarFolderOrContactsFolder()).get(0)).getUnreadCount();

        /*GetFolderType getFolderType = new GetFolderType();
        getFolderType.setFolderShape(folderShape);
        NonEmptyArrayOfBaseFolderIdsType folderArray = new NonEmptyArrayOfBaseFolderIdsType();

        List<BaseFolderType> folders = ((FindFolderResponseMessageType) findFolderResult.value.getResponseMessages().getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage().get(0).getValue()).getRootFolder().getFolders().getFolderOrCalendarFolderOrContactsFolder();

        for (BaseFolderType folder : folders) {
            FolderIdType folderIdType = new FolderIdType();
            folderIdType.setId(folder.getFolderId().getId());
            folderIdType.setChangeKey(folder.getFolderId().getChangeKey());
            folderArray.getFolderIdOrDistinguishedFolderId().add(folderIdType);
        }
                
        getFolderType.setFolderIds(folderArray);

        Holder<GetFolderResponseType> folderResult = new Holder<GetFolderResponseType>();
        exchangeService.getFolder(getFolderType, impersonation, null, null, null, null, folderResult, null);

        List<JAXBElement<? extends ResponseMessageType>> list = folderResult.value.getResponseMessages().getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage();

        for (JAXBElement<? extends ResponseMessageType> jaxbElement : list) {
            ResponseMessageType value = jaxbElement.getValue();
            FolderInfoResponseMessageType info = (FolderInfoResponseMessageType) value;

            System.out.println(info.getFolders().getFolderOrCalendarFolderOrContactsFolder().get(0).getTotalCount());
            System.out.println(jaxbElement.getDeclaredType().toString());
        }*/
    }

    @Test
    @Ignore
    public void testAuth() throws Exception {

        Authenticator.setDefault(new Authenticator() {
            public PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication("vgregion.se\\TK.Exch.Imp.Contact", "WI21aG0O".toCharArray());
            }
        });

        URL url = new URL("https://webmail.vgregion.se/ews/services.wsdl");

        URLConnection urlConnection = url.openConnection(/*new Proxy(Proxy.Type.HTTP, new InetSocketAddress(8888))*/);

        HttpURLConnection httpConnection = (HttpURLConnection) urlConnection;

        InputStream inputStream = httpConnection.getInputStream();

        String response = IOUtils.toString(inputStream);

        System.out.println(response);
    }
}
