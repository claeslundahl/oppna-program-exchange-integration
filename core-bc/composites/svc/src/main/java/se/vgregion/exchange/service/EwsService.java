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
 * Service class for fetching various items and information from Exchange Web Services.
 *
 * @author Patrik Bergström
 */
public class EwsService {

    private static final Logger LOGGER = LoggerFactory.getLogger(EwsService.class);

    private final ObjectFactory objectFactory = new ObjectFactory();
    private ExchangeServicePortType exchangeServicePort;
    private final LdapService ldapService;

    @org.springframework.beans.factory.annotation.Value("${ews.user}")
    private String ewsUser;

    @org.springframework.beans.factory.annotation.Value("${ews.password}")
    private String ewsPassword;

    /**
     * Constructor where the {@link com.microsoft.schemas.exchange.services._2006.messages.ExchangeServicePortType} will
     * be constructed from a wsdl file on the classpath.
     *
     * @param ldapService ldapService
     */
    public EwsService(LdapService ldapService) {
        this.ldapService = ldapService;

        URL wsdlLocation = EwsService.class.getClassLoader().getResource("wsdl/services.wsdl");
        exchangeServicePort = new ExchangeService(wsdlLocation).getExchangeService();
    }

    /**
     * Constructor.
     *
     * @param ldapService ldapService
     * @param exchangeServicePort exchangeServicePort
     */
    public EwsService(LdapService ldapService, ExchangeServicePortType exchangeServicePort) {
        this.ldapService = ldapService;
        this.exchangeServicePort = exchangeServicePort;
    }

    /**
     * Initialization of authentication for NTLM which is needed for the Exchange web service.
     */
    @PostConstruct
    public void init() {

        LOGGER.info("Initializing authentication.");

        // Make NTLM work
        Authenticator.setDefault(new Authenticator() {
            public PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(ewsUser, ewsPassword.toCharArray());
            }
        });

        Client clientProxy = ClientProxy.getClient(exchangeServicePort);
        HTTPConduit conduit = (HTTPConduit) clientProxy.getConduit();

        HTTPClientPolicy client = new HTTPClientPolicy();

        // These are needed to make NTLM authentication work.
        client.setAllowChunking(false);
        client.setAutoRedirect(true);

        conduit.setClient(client);
    }

    /**
     * Fetch unread emails from a user's inbox folder (named "Inbox" or "Inkorg") with all common properties but
     * without message body.
     *
     * @param userId the user id
     * @param maxNumber the maximum number of unread emails to retrieve
     * @return a list of unread emails
     */
    public List<MessageType> fetchUnreadEmails(String userId, int maxNumber) {

        // First fetch the inbox folder
        FolderType inboxFolder = findInboxFolder(userId);

        if (inboxFolder == null) {
            return null;
        }

        ConstantValueType constantValueType = objectFactory.createConstantValueType();
        constantValueType.setValue("0");

        PathToUnindexedFieldType fieldType = new PathToUnindexedFieldType();
        fieldType.setFieldURI(UnindexedFieldURIType.MESSAGE_IS_READ);
        PathToUnindexedFieldType pathToUnindexedFieldType = objectFactory.createPathToUnindexedFieldType();
        pathToUnindexedFieldType.setFieldURI(UnindexedFieldURIType.MESSAGE_IS_READ);

        FieldURIOrConstantType fieldURIOrConstant = objectFactory.createFieldURIOrConstantType();
        fieldURIOrConstant.setConstant(constantValueType);

        IsEqualToType filterMessages = new IsEqualToType();
        filterMessages.setFieldURIOrConstant(fieldURIOrConstant);
        filterMessages.setPath(objectFactory.createFieldURI(pathToUnindexedFieldType));

        JAXBElement<IsEqualToType> messageReadEqualsFalse = objectFactory.createIsEqualTo(filterMessages);

        RestrictionType restriction = new RestrictionType();
        restriction.setSearchExpression(messageReadEqualsFalse);

        NonEmptyArrayOfBaseFolderIdsType parentFolderIds = objectFactory.createNonEmptyArrayOfBaseFolderIdsType();
        parentFolderIds.getFolderIdOrDistinguishedFolderId().add(inboxFolder.getFolderId());

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

        if (impersonation == null) {
            return null;
        }

        exchangeServicePort.findItem(findItemType, impersonation, null, null, null, null, null,
                findItemResult, null);

        List<JAXBElement<? extends ResponseMessageType>> list = findItemResult.value.getResponseMessages()
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage();
        FindItemParentType rootFolder = ((FindItemResponseMessageType) list.get(0).getValue()).getRootFolder();
        List<? extends ItemType> emails = rootFolder.getItems().getItemOrMessageOrCalendarItem();

        // To fetch message bodies
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

    /**
     * Fetch all calendar events for a user for a given time period.
     *
     * @param userId the user id
     * @param startDate startDate
     * @param endDate endDate
     * @return all calendar events for a user for the given time period
     */
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

        if (exchangeImpersonation == null) {
            return null;
        }

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

    /**
     * Fetch the number of unread emails in the user's inbox.
     *
     * @param userId the user id
     * @return the number of unread emails in the user's inbox
     */
    public Integer fetchInboxUnreadCount(String userId) {

        FolderType inbox = findInboxFolder(userId);

        if (inbox == null) {
            return null;
        }

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

        ConstantValueType c2 = new ConstantValueType();
        c2.setValue("Inbox");

        FieldURIOrConstantType displayNameField = new FieldURIOrConstantType();
        displayNameField.setConstant(c1);

        FieldURIOrConstantType displayNameField2 = new FieldURIOrConstantType();
        displayNameField2.setConstant(c2);

        PathToUnindexedFieldType fieldType = new PathToUnindexedFieldType();
        fieldType.setFieldURI(UnindexedFieldURIType.FOLDER_DISPLAY_NAME);

        IsEqualToType equalToType = new IsEqualToType();
        equalToType.setFieldURIOrConstant(displayNameField);
        equalToType.setPath(new JAXBElement<BasePathToElementType>(
                new QName("http://schemas.microsoft.com/exchange/services/2006/types", "FieldURI"),
                BasePathToElementType.class, fieldType));

        IsEqualToType equalToType2 = new IsEqualToType();
        equalToType2.setFieldURIOrConstant(displayNameField2);
        equalToType2.setPath(new JAXBElement<BasePathToElementType>(
                new QName("http://schemas.microsoft.com/exchange/services/2006/types", "FieldURI"),
                BasePathToElementType.class, fieldType));

        MultipleOperandBooleanExpressionType orExpression = new OrType();
        orExpression.getSearchExpression().add(objectFactory.createIsEqualTo(equalToType));
        orExpression.getSearchExpression().add(objectFactory.createIsEqualTo(equalToType2));

        RestrictionType restriction = new RestrictionType();
        restriction.setSearchExpression(objectFactory.createSearchExpression(orExpression));

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

        if (impersonation == null) {
            return null;
        }

        exchangeServicePort.findFolder(findFolderType, impersonation, null, null,
                null, null, findFolderResult, null);

        ArrayOfResponseMessagesType responseMessages = findFolderResult.value.getResponseMessages();

        List<JAXBElement<? extends ResponseMessageType>> list = responseMessages
                .getCreateItemResponseMessageOrDeleteItemResponseMessageOrGetItemResponseMessage();

        if (list == null || list.size() == 0) {
            return null;
        }

        JAXBElement<? extends ResponseMessageType> jaxbElement = list.get(0);

        return (FolderType) (((((FindFolderResponseMessageType) jaxbElement.getValue()).getRootFolder())
                .getFolders()).getFolderOrCalendarFolderOrContactsFolder()).get(0);
    }

    private ExchangeImpersonationType getExchangeImpersonation(String userId) {
        String userSid = fetchUserSid(userId);

        if (userSid == null) {
            return null;
        }

        ConnectingSIDType connectingSID = new ConnectingSIDType();
        connectingSID.setSID(userSid);

        ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
        impersonation.setConnectingSID(connectingSID);
        return impersonation;
    }

    String fetchUserSid(String userId) {
        LdapUser[] ldapUser = ldapService.search("", String.format("(&(objectClass=person)(cn=%s))", userId));

        if (ldapUser == null || ldapUser.length == 0) {
            return null;
        }

        ArrayList attributes = ldapUser[0].getAttributes().get("objectSid");
        String sid = LdapUtils.convertBinarySidToString((byte[]) attributes.get(0));

        return sid;
    }
}
