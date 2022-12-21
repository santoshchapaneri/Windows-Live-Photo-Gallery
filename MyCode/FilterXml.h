#pragma once

#include <msxml6.h>

namespace FilterXml
{

class FilterXml
{
public:

    // Loading/Saving
    static void LoadXmlDocument(CString strFileName, IXMLDOMDocument** ppXmlDoc) throw(...);
    static void CreateXmlDocument(IXMLDOMDocument** ppXmlDoc) throw(...);
    static void SaveXmlDocument(CString strOutputPath, IXMLDOMDocument* pXmlDoc) throw(...);

    // Node Interaction
    static CComBSTR GetAttributeValue(CComBSTR strAttributeName, IXMLDOMNode* pNode) throw(...);
    static bool SetAttributeValue(CComBSTR strAttributeName, CComBSTR strAttributeValue, IXMLDOMNode* pNode) throw(...);
    static void CreateNode(IXMLDOMDocument* pXmlDocument, CComBSTR strNodeName, IXMLDOMNode** ppNode) throw(...);
    static void CopyNodeAttributes(IXMLDOMNode* pFromNode, IXMLDOMNode* pToNode) throw(...);
    static bool CopyNodeAttribute(CComBSTR strAttributeName, IXMLDOMNode* pFromNode, IXMLDOMNode* pToNode) throw(...); // Returns true/false depending on whether or not the requested attribute was found in the "From" node
    static void AddAttributeToNode(CComBSTR strAttributeName, CComBSTR strAttributeValue, IXMLDOMNode* pNode) throw(...);
    static void SetNodeText(CComBSTR strNodeText, IXMLDOMNode* pNode) throw(...);
    static CComBSTR GetSingleNodeText(CComBSTR strNodeName, IXMLDOMDocument* pXmlDoc) throw(...);

protected:

private:
    FilterXml();
    FilterXml(FilterXml&);
    ~FilterXml();
};

}

