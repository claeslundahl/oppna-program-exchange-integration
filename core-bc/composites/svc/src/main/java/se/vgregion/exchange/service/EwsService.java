package se.vgregion.exchange.service;

import com.microsoft.schemas.exchange.services._2006.messages.*;
import com.microsoft.schemas.exchange.services._2006.types.*;
import com.microsoft.schemas.exchange.services._2006.types.ObjectFactory;
import org.apache.cxf.endpoint.Client;
import org.apache.cxf.frontend.ClientProxy;
import org.apache.cxf.transport.http.HTTPConduit;
import org.apache.cxf.transports.http.configuration.HTTPClientPolicy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.ldap.support.LdapUtils;
import se.vgregion.ldapservice.LdapService;
import se.vgregion.ldapservice.LdapUser;

import javax.annotation.PostConstruct;
import javax.xml.bind.JAXBElement;
import javax.xml.datatype.DatatypeConfigurationException;
import javax.xml.datatype.DatatypeFactory;
import javax.xml.namespace.QName;
import javax.xml.ws.Holder;
import java.net.Authenticator;
import java.net.PasswordAuthentication;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;

/**
 * @author Patrik Bergström
 */
public class EwsService {

    private static final Logger LOGGER = LoggerFactory.getLogger(EwsService.class);

    private final ObjectFactory objectFactory = new ObjectFactory();
    private ExchangeServicePortType exchangeServicePort;
    private final LdapService ldapService;
    private final URL wsdlLocation;

    @org.springframework.beans.factory.annotation.Value("${ews.user}")
    private String ewsUser;

    @org.springframework.beans.factory.annotation.Value("${ews.password}")
    private String ewsPassword;

    public EwsService(LdapService ldapService) {
        this.ldapService = ldapService;

        this.wsdlLocation = EwsService.class.getClassLoader().getResource("wsdl/services.wsdl");
        exchangeServicePort = new ExchangeService(wsdlLocation).getExchangeService();
    }

    @PostConstruct
    public void init() {
        // Make NTLM work
        Authenticator.setDefault(new Authenticator() {
            public PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(ewsUser, ewsPassword.toCharArray());
            }
        });

        Client clientProxy = ClientProxy.getClient(exchangeServicePort);
//        clientProxy.getOutInterceptors().add(new LoggingOutInterceptor());
//        clientProxy.getInInterceptors().add(new LoggingInInterceptor());
//        clientProxy.getInFaultInterceptors().add(new LoggingInInterceptor());
        HTTPConduit conduit = (HTTPConduit) clientProxy.getConduit();

        HTTPClientPolicy client = new HTTPClientPolicy();

        // These are needed to make NTLM authentication work.
        client.setAllowChunking(false);
        client.setAutoRedirect(true);

//        client.setProxyServer("127.0.0.1");
//        client.setProxyServerPort(8888);

        conduit.setClient(client);
    }

    public List<MessageType> fetchUnreadEmails(String userId, int maxNumber) {

        ConstantValueType constantValueType = objectFactory.createConstantValueType();
        constantValueType.setValue("0");

        PathToUnindexedFieldType fieldType = new PathToUnindexedFieldType();
        fieldType.setFieldURI(UnindexedFieldURIType.MESSAGE_IS_READ);
        PathToUnindexedFieldType pathToUnindexedFieldType = objectFactory.createPathToUnindexedFieldType();
        pathToUnindexedFieldType.setFieldURI(UnindexedFieldURIType.MESSAGE_IS_READ);

        FieldURIOrConstantType fieldURIOrConstant = objectFactory.createFieldURIOrConstantType();
        fieldURIOrConstant.setConstant(constantValueType);
//        fieldURIOrConstant.setPath(objectFactory.createFieldURI(pathToUnindexedFieldType));
        
        IsEqualToType filterMessages = new IsEqualToType();
        filterMessages.setFieldURIOrConstant(fieldURIOrConstant);
        filterMessages.setPath(objectFactory.createFieldURI(pathToUnindexedFieldType));

        JAXBElement<IsEqualToType> messageReadEqualsFalse = objectFactory.createIsEqualTo(filterMessages);

        RestrictionType restriction = new RestrictionType();
        restriction.setSearchExpression(messageReadEqualsFalse);

        NonEmptyArrayOfBaseFolderIdsType parentFolderIds = objectFactory.createNonEmptyArrayOfBaseFolderIdsType();
        parentFolderIds.getFolderIdOrDistinguishedFolderId().add(findInboxFolder(userId).getFolderId());

        IndexedPageViewType indexedPageViewType = objectFactory.createIndexedPageViewType();
        indexedPageViewType.setBasePoint(IndexBasePointType.BEGINNING);
        indexedPageViewType.setOffset(0);
        indexedPageViewType.setMaxEntriesReturned(maxNumber);

        ItemResponseShapeType itemResponseShapeType = objectFactory.createItemResponseShapeType();
        itemResponseShapeType.setBaseShape(DefaultShapeNamesType.ALL_PROPERTIES);

        FindItemType findItemType = new FindItemType();
        findItemType.setRestriction(restriction);
        findItemType.setParentFolderIds(parentFolderIds);
        findItemType.setTraversal(ItemQueryTraversalType.SHALLOW);
        findItemType.setIndexedPageItemView(indexedPageViewType);
        findItemType.setItemShape(itemResponseShapeType);

        Holder<FindItemResponseType> findItemResult = new Holder<FindItemResponseType>();

        ExchangeImpersonationType impersonation = getExchangeImpersonation(userId);
        exchangeServicePort.findItem(findItemType, impersonation, null, null, null, null, null,
                findItemResult, null);

        List<JAXBElement<? extends ResponseMessageType>> list = findItemResult.value.getResponseMessages()
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage();
        FindItemParentType rootFolder = ((FindItemResponseMessageType) list.get(0).getValue()).getRootFolder();
        List<? extends ItemType> emails = rootFolder.getItems().getItemOrMessageOrCalendarItem();

        /*for (ItemType email : emails) {
            NonEmptyArrayOfBaseItemIdsType ids = objectFactory.createNonEmptyArrayOfBaseItemIdsType();
            ids.getItemIdOrOccurrenceItemIdOrRecurringMasterItemId().add(email.getItemId());

            ItemResponseShapeType shapeType = objectFactory.createItemResponseShapeType();
            shapeType.setBaseShape(DefaultShapeNamesType.ALL_PROPERTIES);
            shapeType.setBodyType(BodyTypeResponseType.TEXT);

            GetItemType getItemType = new GetItemType();
            getItemType.setItemIds(ids);
            getItemType.setItemShape(shapeType);

            Holder<GetItemResponseType> getItemResult = new Holder<GetItemResponseType>();

            exchangeServicePort.getItem(getItemType, impersonation, null, null, null, null, null, getItemResult, null);

            ItemInfoResponseMessageType messageType = (ItemInfoResponseMessageType) getItemResult.value
                    .getResponseMessages()
                    .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage()
                    .get(0).getValue();

            BodyType body = messageType.getItems().getItemOrMessageOrCalendarItem().get(0).getBody();
            email.setBody(body);
        }*/

        return (List<MessageType>) emails;
    }

    public List<CalendarItemType> fetchCalendarEvents(String userId, Date startDate, Date endDate) {
        GregorianCalendar startDateCalendar = new GregorianCalendar();
        startDateCalendar.setTime(startDate);

        GregorianCalendar endDateCalendar = new GregorianCalendar();
        endDateCalendar.setTime(endDate);

        CalendarViewType calendarView = new CalendarViewType();
        DatatypeFactory datatypeFactory;
        try {
            datatypeFactory = DatatypeFactory.newInstance();
        } catch (DatatypeConfigurationException e) {
            throw new RuntimeException(e);
        }

        calendarView.setStartDate(datatypeFactory.newXMLGregorianCalendar(startDateCalendar));
        calendarView.setEndDate(datatypeFactory.newXMLGregorianCalendar(endDateCalendar));

        ItemResponseShapeType shape = new ItemResponseShapeType();
        shape.setBaseShape(DefaultShapeNamesType.ALL_PROPERTIES);

        DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
        folderId.setId(DistinguishedFolderIdNameType.CALENDAR);

        NonEmptyArrayOfBaseFolderIdsType parentFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
        parentFolderIds.getFolderIdOrDistinguishedFolderId().add(folderId);

        FindItemType findItem = new FindItemType();
        findItem.setCalendarView(calendarView);
        findItem.setItemShape(shape);
        findItem.setParentFolderIds(parentFolderIds);
        findItem.setTraversal(ItemQueryTraversalType.SHALLOW);

        ExchangeImpersonationType exchangeImpersonation = getExchangeImpersonation(userId);

        Holder<FindItemResponseType> findItemResult = new Holder<FindItemResponseType>();

        exchangeServicePort.findItem(findItem, exchangeImpersonation, null, null, null, null, null, findItemResult,
                null);

        List<JAXBElement<? extends ResponseMessageType>> responseMessage = findItemResult.value.getResponseMessages()
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage();

        if (responseMessage.size() == 0) {
            return null;
        } else if (responseMessage.size() > 1) {
            throw new RuntimeException("Expected only one responseMessage.");
        }

        JAXBElement<? extends ResponseMessageType> jaxbElement = responseMessage.get(0);

        List<? extends ItemType> items = ((FindItemResponseMessageType) jaxbElement.getValue()).getRootFolder()
                .getItems().getItemOrMessageOrCalendarItem();

        return (List<CalendarItemType>) items;
    }

    public Integer fetchInboxUnreadCount(String userId) {

        FolderType inbox = findInboxFolder(userId);

        return inbox.getUnreadCount();
    }

    private FolderType findInboxFolder(String userId) {
        Holder<FindFolderResponseType> findFolderResult = new Holder<FindFolderResponseType>();

        IndexedPageViewType indexedPageViewType = new IndexedPageViewType();
        indexedPageViewType.setBasePoint(IndexBasePointType.BEGINNING);
        indexedPageViewType.setOffset(0);
        indexedPageViewType.setMaxEntriesReturned(1);

        ConstantValueType c1 = new ConstantValueType();
        c1.setValue("Inkorg");

        FieldURIOrConstantType displayNameField = new FieldURIOrConstantType();
        displayNameField.setConstant(c1);

        PathToUnindexedFieldType fieldType = new PathToUnindexedFieldType();
        fieldType.setFieldURI(UnindexedFieldURIType.FOLDER_DISPLAY_NAME);

        IsEqualToType equalToType = new IsEqualToType();
        equalToType.setFieldURIOrConstant(displayNameField);
        equalToType.setPath(new JAXBElement<BasePathToElementType>(
                new QName("http://schemas.microsoft.com/exchange/services/2006/types", "FieldURI"),
                BasePathToElementType.class, fieldType));

        QName qName = new QName("http://schemas.microsoft.com/exchange/services/2006/types", "IsEqualTo");

        RestrictionType restriction = new RestrictionType();
        restriction.setSearchExpression(new JAXBElement<SearchExpressionType>(qName, SearchExpressionType.class,
                equalToType));

        DistinguishedFolderIdType distinguishedFolderIdType = new DistinguishedFolderIdType();
        distinguishedFolderIdType.setId(DistinguishedFolderIdNameType.ROOT);

        NonEmptyArrayOfBaseFolderIdsType folderIdsType = new NonEmptyArrayOfBaseFolderIdsType();
        folderIdsType.getFolderIdOrDistinguishedFolderId().add(distinguishedFolderIdType);

        FolderResponseShapeType folderShape = new FolderResponseShapeType();
        folderShape.setBaseShape(DefaultShapeNamesType.ALL_PROPERTIES);

        FindFolderType findFolderType = new FindFolderType();
        findFolderType.setTraversal(FolderQueryTraversalType.DEEP);
        findFolderType.setIndexedPageFolderView(indexedPageViewType);
        findFolderType.setRestriction(restriction);
        findFolderType.setParentFolderIds(folderIdsType);
        findFolderType.setFolderShape(folderShape);

        ExchangeImpersonationType impersonation = getExchangeImpersonation(userId);

        exchangeServicePort.findFolder(findFolderType, impersonation, null, null,
                null, null, findFolderResult, null);

        ArrayOfResponseMessagesType responseMessages = findFolderResult.value.getResponseMessages();

        JAXBElement<? extends ResponseMessageType> jaxbElement = responseMessages
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage().get(0);

        return (FolderType) (((((FindFolderResponseMessageType) jaxbElement.getValue()).getRootFolder())
                .getFolders()).getFolderOrCalendarFolderOrContactsFolder()).get(0);
    }

    private ExchangeImpersonationType getExchangeImpersonation(String userId) {
        String userSid = fetchUserSid(userId);

        ConnectingSIDType connectingSID = new ConnectingSIDType();
        connectingSID.setSID(userSid);

        ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
        impersonation.setConnectingSID(connectingSID);
        return impersonation;
    }

    String fetchUserSid(String userId) {
        LdapUser[] ldapUser = ldapService.search("", String.format("(&(objectClass=person)(cn=%s))", userId));

        if (ldapUser.length != 1) {
            throw new RuntimeException("Expected exactly one match. " + userId + " resulted in " + ldapUser.length
                    + " matches.");
        }

        ArrayList attributes = ldapUser[0].getAttributes().get("objectSid");
        String sid = LdapUtils.convertBinarySidToString((byte[]) attributes.get(0));

        return sid;
    }
}
