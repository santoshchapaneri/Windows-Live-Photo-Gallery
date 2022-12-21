#include "FilterXml.h"

namespace FilterXml
{

void FilterXml::LoadXmlDocument(CString strFileName, IXMLDOMDocument** ppXmlDoc)
{
    VERIFY_POINTER_OR_THROW(ppXmlDoc, L"Invalid pointer passed into function.");

    CComPtr<IXMLDOMDocument2> spXmlDoc;
    VERIFY_HR_OR_THROW(CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&spXmlDoc)), "Failed to create the IXMLDOMDocument");

    // Open the XML file
    VARIANT vtFile;
    CComBSTR bstrFile(strFileName);
    VARIANT_BOOL vbOpen;

    V_VT(&vtFile) = VT_BSTR;
    V_BSTR(&vtFile) = bstrFile;

    VERIFY_HR_OR_THROW(spXmlDoc->load(vtFile, &vbOpen), "Failed to load the xml document.");
    VERIFY_OR_THROW(!!vbOpen, "Failed to load the xml document.");

    spXmlDoc.CopyTo(ppXmlDoc);
}

void FilterXml::CreateXmlDocument(IXMLDOMDocument** ppXmlDoc)
{
    VERIFY_POINTER_OR_THROW(ppXmlDoc, L"Invalid pointer passed into function.");

    CComPtr<IXMLDOMDocument2> spXmlDoc;
    VERIFY_HR_OR_THROW(CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&spXmlDoc)), "Failed to create the IXMLDOMDocument");

    spXmlDoc.CopyTo(ppXmlDoc);
}

void FilterXml::SaveXmlDocument(CString strOutputPath, IXMLDOMDocument* pXmlDoc)
{
    VERIFY_POINTER_OR_THROW(pXmlDoc, L"Invalid pointer passed into function.");

    // Save the XML document
    VARIANT vtFile;
    CComBSTR bstrFile(strOutputPath);
    V_VT(&vtFile) = VT_BSTR;
    V_BSTR(&vtFile) = bstrFile;

    VERIFY_HR_OR_THROW(pXmlDoc->save(vtFile), "Failed to save the new file");
}

void FilterXml::CopyNodeAttributes(IXMLDOMNode* pFromNode, IXMLDOMNode* pToNode)
{
    VERIFY_POINTER_OR_THROW(pFromNode, L"Invalid pointer passed into function.");
    VERIFY_POINTER_OR_THROW(pToNode, L"Invalid pointer passed into function.");

    // Get the attributes from the "From" node
    CComPtr<IXMLDOMNamedNodeMap> spAttributes;
    VERIFY_HR_OR_THROW(pFromNode->get_attributes(&spAttributes), "Failed to get the attributes from the input node");

    // Copy them all to the "To" node
    CComPtr<IXMLDOMNode> spAttribute;
    spAttributes->nextNode(&spAttribute);
    while(spAttribute != NULL)
    {
        CComBSTR bstrName;
        CComBSTR bstrValue;

        VERIFY_HR_OR_THROW(spAttribute->get_nodeName(&bstrName), "Failed to get the name of the attribute to copy");
        VERIFY_HR_OR_THROW(spAttribute->get_text(&bstrValue), "Failed to get the value of the attribute to copy");

        AddAttributeToNode(bstrName.m_str, bstrValue.m_str, pToNode);

        // Release and move on to the next Attribute
        spAttribute.Release();
        spAttributes->nextNode(&spAttribute);
    }
}

bool FilterXml::CopyNodeAttribute(CComBSTR strAttributeName, IXMLDOMNode* pFromNode, IXMLDOMNode* pToNode)
{
    VERIFY_POINTER_OR_THROW(pFromNode, L"Invalid pointer passed into function.");
    VERIFY_POINTER_OR_THROW(pToNode, L"Invalid pointer passed into function.");

    // Get the attributes from the "From" node
    CComPtr<IXMLDOMNamedNodeMap> spAttributes;
    VERIFY_HR_OR_THROW(pFromNode->get_attributes(&spAttributes), "Failed to get the attributes from the input node");

    CComPtr<IXMLDOMNode> spAttribute;
    bool fFound = (spAttributes->getNamedItem(strAttributeName, &spAttribute) == S_OK);
    if(fFound)
    {
        CComBSTR bstrValue;
        spAttribute->get_text(&bstrValue);
        AddAttributeToNode(strAttributeName, bstrValue.m_str, pToNode);
    }
    
    return fFound;
}

void FilterXml::CreateNode(IXMLDOMDocument* pXmlDocument, CComBSTR strNodeName, IXMLDOMNode** ppNode)
{
    VERIFY_POINTER_OR_THROW(pXmlDocument, L"Invalid pointer passed into function.");
    VERIFY_POINTER_OR_THROW(ppNode, L"Invalid pointer passed into function.");

    // Create an element with the specified name
    CComPtr<IXMLDOMElement> spElement;
    VERIFY_HR_OR_THROW(pXmlDocument->createElement(strNodeName, &spElement), "Failed to create the specified XML Element");

    // QI for IXMLDOMNode and copy to the out param
    CComQIPtr<IXMLDOMNode> spNode(spElement);
    spNode.CopyTo(ppNode);
}

void FilterXml::AddAttributeToNode(CComBSTR strAttributeName, CComBSTR strAttributeValue, IXMLDOMNode* pNode)
{
    VERIFY_POINTER_OR_THROW(pNode, L"Invalid pointer passed into function.");

    // Get the Owner Document of the "To" node so we can creat new attributes for it
    CComPtr<IXMLDOMDocument> spToXmlDoc;
    VERIFY_HR_OR_THROW(pNode->get_ownerDocument(&spToXmlDoc), "Failed to get the owner document the node");

    // Create the new attribute and copy it to the node
    CComPtr<IXMLDOMAttribute> spNewAttribute;
    VERIFY_HR_OR_THROW(spToXmlDoc->createAttribute(strAttributeName, &spNewAttribute), "Failed to create the attribute");
    VERIFY_HR_OR_THROW(spNewAttribute->put_text(strAttributeValue), "Failed to set the text for attribute");

    CComQIPtr<IXMLDOMElement> spToElement(pNode);
    VERIFY_POINTER_OR_THROW(spToElement, "Failed to QI the IXMLDOMNode for IXMLDOMElement");
    VERIFY_HR_OR_THROW(spToElement->setAttributeNode(spNewAttribute, NULL), "Failed to set the attribute");  
}

CComBSTR FilterXml::GetAttributeValue(CComBSTR strAttributeName, IXMLDOMNode* pNode)
{
    VERIFY_POINTER_OR_THROW(pNode, L"Invalid pointer passed into function.");
    CComBSTR bstrAttributeText;

    CComPtr<IXMLDOMNamedNodeMap> spNodeAttributes;
    HRESULT hr = pNode->get_attributes(&spNodeAttributes);

    if(S_OK == hr)
    {
        CComPtr<IXMLDOMNode> spAttribute;
        
        // Look for the named item
        hr = spNodeAttributes->getNamedItem(strAttributeName, &spAttribute);
        if(S_OK == hr)
        {
            // Query for the actual text of the attribute
            spAttribute->get_text(&bstrAttributeText);
        }
    }

    return bstrAttributeText;
}

bool FilterXml::SetAttributeValue(CComBSTR strAttributeName, CComBSTR strAttributeValue, IXMLDOMNode* pNode)
{
    VERIFY_POINTER_OR_THROW(pNode, L"Invalid pointer passed into function.");

    bool fFound = false;

    CComPtr<IXMLDOMNamedNodeMap> spNodeAttributes;
    VERIFY_HR_OR_THROW(pNode->get_attributes(&spNodeAttributes), "Failed to get the attributes for the Node");

    CComPtr<IXMLDOMNode> spAttribute;
    
    // Look for the named item
    if(S_OK == spNodeAttributes->getNamedItem(strAttributeName, &spAttribute))
    {
        // Set the text of the attribute
        VERIFY_HR_OR_THROW(spAttribute->put_text(strAttributeValue), "Failed to set the text for the specified attribute");
        fFound = true;
    }
    
    return fFound;
}

void FilterXml::SetNodeText(CComBSTR strNodeText, IXMLDOMNode* pNode)
{
    VERIFY_HR_OR_THROW(pNode->put_text(strNodeText), "Failed to set the text for the node");
}

CComBSTR FilterXml::GetSingleNodeText(CComBSTR strNodeName, IXMLDOMDocument* pXmlDoc)
{
    VERIFY_POINTER_OR_THROW(pXmlDoc, L"Invalid pointer passed into function.");
    CComBSTR bstrLocation;

    // Get list of nodes that match given name
    CComPtr<IXMLDOMNodeList> spNodeList;
    HRESULT hr = pXmlDoc->getElementsByTagName(strNodeName, &spNodeList);
    
    if (S_OK == hr)
    {
        // Get one and only node
        LONG lLength;
        VERIFY_HR_OR_THROW(spNodeList->get_length(&lLength), L"Couldn't get node length from config file.");
        VERIFY_OR_THROW(lLength == 1, L"Didn't get single node.");

        CComPtr<IXMLDOMNode> spNode;
        hr = spNodeList->get_item(0, &spNode);
        if (S_OK == hr)
        {
            hr = spNode->get_text(&bstrLocation);
        }
    }

    return bstrLocation;
}

}