// ----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// ----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
// ----------------------------------------------------------------------------------

//
// <copyright file="Authentication.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Collections.Specialized;
using System.Net;
using System.Web;
using System.Security.Cryptography;
using System.Globalization;
using System.Diagnostics;


namespace Microsoft.Samples.ServiceHosting.StorageClient
{
    /// <summary>
    /// This type represents the different constituent parts that make up a resource Uri in the context of cloud services.
    /// </summary>
    public class ResourceUriComponents
    {
        /// <summary>
        /// The account name in the URI.
        /// </summary>
        public string AccountName { get; set; }

        /// <summary>
        /// This is really the first component (delimited by '/') after the account name. Since it happens to
        /// be a container name in the context of all our storage services (containers in blob storage,
        /// queues in the queue service and table names in table storage), it's named as ContainerName to make it more 
        /// readable at the cost of slightly being incorrectly named.
        /// </summary>
        public string ContainerName { get; set; }

        /// <summary>
        /// The remaining string in the URI.
        /// </summary>
        public string RemainingPart { get; set; }

        /// <summary>
        /// Construct a ResourceUriComponents object.
        /// </summary>
        /// <param name="accountName">The account name that should become part of the URI.</param>
        /// <param name="containerName">The container name (container, queue or table name) that should become part of the URI.</param>
        /// <param name="remainingPart">Remaining part of the URI.</param>
        public ResourceUriComponents(string accountName, string containerName, string remainingPart)
        {
            this.AccountName = accountName;
            this.ContainerName = containerName;
            this.RemainingPart = remainingPart;
        }

        /// <summary>
        /// Construct a ResourceUriComponents object.        
        /// </summary>
        /// <param name="accountName">The account name that should become part of the URI.</param>
        /// <param name="containerName">The container name (container, queue or table name) that should become part of the URI.</param>
        public ResourceUriComponents(string accountName, string containerName)
            : this(accountName, containerName, null)
        {
        }

        /// <summary>
        /// Construct a ResourceUriComponents object.        
        /// </summary>
        /// <param name="accountName">The account name that should become part of the URI.</param>
        public ResourceUriComponents(string accountName)
            : this(accountName, null, null)
        {
        }

        /// <summary>
        /// Construct a ResourceUriComponents object.        
        /// </summary>
        public ResourceUriComponents()
        {
        }
    }

    internal static class MessageCanonicalizer
    {
        /// <summary>
        /// An internal class that stores the canonicalized string version of an HTTP request.
        /// </summary>
        private class CanonicalizedString
        {
            private StringBuilder canonicalizedString = new StringBuilder();

            /// <summary>
            /// Property for the canonicalized string.
            /// </summary>
            internal string Value
            {
                get
                {
                    return canonicalizedString.ToString();
                }
            }

            /// <summary>
            /// Constructor for the class.
            /// </summary>
            /// <param name="initialElement">The first canonicalized element to start the string with.</param>
            internal CanonicalizedString(string initialElement)
            {
                canonicalizedString.Append(initialElement);
            }

            /// <summary>
            /// Append additional canonicalized element to the string.
            /// </summary>
            /// <param name="element">An additional canonicalized element to append to the string.</param>
            internal void AppendCanonicalizedElement(string element)
            {
                canonicalizedString.Append(StorageHttpConstants.ConstChars.Linefeed);
                canonicalizedString.Append(element);
            }
        }

        /// <summary>
        /// Create a canonicalized string out of HTTP request header contents for signing 
        /// blob/queue requests with the Shared Authentication scheme. 
        /// </summary>
        /// <param name="address">The uri address of the HTTP request.</param>
        /// <param name="uriComponents">Components of the Uri extracted out of the request.</param>
        /// <param name="method">The method of the HTTP request (GET/PUT, etc.).</param>
        /// <param name="contentType">The content type of the HTTP request.</param>
        /// <param name="date">The date of the HTTP request.</param>
        /// <param name="headers">Should contain other headers of the HTTP request.</param>
        /// <returns>A canonicalized string of the HTTP request.</returns>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1308:NormalizeStringsToUppercase",
            Justification = "Authentication algorithm requires canonicalization by converting to lower case")]
        internal static string CanonicalizeHttpRequest(
            Uri address, 
            ResourceUriComponents uriComponents, 
            string method, 
            string contentType, 
            string date, 
            NameValueCollection headers)
        {
            // The first element should be the Method of the request.
            // I.e. GET, POST, PUT, or HEAD.
            CanonicalizedString canonicalizedString = new CanonicalizedString(method);

            // The second element should be the MD5 value.
            // This is optional and may be empty.
            string httpContentMD5Value = string.Empty;

            // First extract all the content MD5 values from the header.
            ArrayList httpContentMD5Values = HttpRequestAccessor.GetHeaderValues(headers, StorageHttpConstants.HeaderNames.ContentMD5);

            // If we only have one, then set it to the value we want to append to the canonicalized string.
            if (httpContentMD5Values.Count == 1)
            {
                httpContentMD5Value = (string)httpContentMD5Values[0];
            }

            canonicalizedString.AppendCanonicalizedElement(httpContentMD5Value);

            // The third element should be the content type.
            canonicalizedString.AppendCanonicalizedElement(contentType);

            // The fourth element should be the request date.
            // See if there's an storage date header.
            // If there's one, then don't use the date header.
            ArrayList httpStorageDateValues = HttpRequestAccessor.GetHeaderValues(headers, StorageHttpConstants.HeaderNames.StorageDateTime);
            if (httpStorageDateValues.Count > 0)
            {
                date = null;
            }

            canonicalizedString.AppendCanonicalizedElement(date);

            // Look for header names that start with StorageHttpConstants.HeaderNames.PrefixForStorageHeader
            // Then sort them in case-insensitive manner.
            ArrayList httpStorageHeaderNameArray = new ArrayList();
            foreach (string key in headers.Keys)
            {
                if (key.ToLowerInvariant().StartsWith(StorageHttpConstants.HeaderNames.PrefixForStorageHeader, StringComparison.Ordinal))
                {
                    httpStorageHeaderNameArray.Add(key.ToLowerInvariant());
                }
            }

            httpStorageHeaderNameArray.Sort();

            // Now go through each header's values in the sorted order and append them to the canonicalized string.
            foreach (string key in httpStorageHeaderNameArray)
            {
                StringBuilder canonicalizedElement = new StringBuilder(key);
                string delimiter = ":";
                ArrayList values = HttpRequestAccessor.GetHeaderValues(headers, key);

                // Go through values, unfold them, and then append them to the canonicalized element string.
                foreach (string value in values)
                {
                    // Unfolding is simply removal of CRLF.
                    string unfoldedValue = value.Replace(StorageHttpConstants.ConstChars.CarriageReturnLinefeed, string.Empty);

                    // Append it to the canonicalized element string.
                    canonicalizedElement.Append(delimiter);
                    canonicalizedElement.Append(unfoldedValue);
                    delimiter = ",";
                }

                // Now, add this canonicalized element to the canonicalized header string.
                canonicalizedString.AppendCanonicalizedElement(canonicalizedElement.ToString());
            }

            // Now we append the canonicalized resource element.
            string canonicalizedResource = GetCanonicalizedResource(address, uriComponents);
            canonicalizedString.AppendCanonicalizedElement(canonicalizedResource);

            return canonicalizedString.Value;
        }

        internal static string GetCanonicalizedResource(Uri address, ResourceUriComponents uriComponents)
        {
            // Algorithem is as follows
            // 1. Start with the empty string ("")
            // 2. Append the account name owning the resource preceded by a /. This is not 
            //    the name of the account making the request but the account that owns the 
            //    resource being accessed.
            // 3. Append the path part of the un-decoded HTTP Request-URI, up-to but not 
            //    including the query string.
            // 4. If the request addresses a particular component of a resource, like?comp=
            //    metadata then append the sub-resource including question mark (like ?comp=
            //    metadata)
            StringBuilder canonicalizedResource = new StringBuilder(StorageHttpConstants.ConstChars.Slash);
            canonicalizedResource.Append(uriComponents.AccountName);

            // Note that AbsolutePath starts with a '/'.
            canonicalizedResource.Append(address.AbsolutePath);

            NameValueCollection queryVariables = HttpUtility.ParseQueryString(address.Query);
            string compQueryParameterValue = queryVariables[StorageHttpConstants.QueryParams.QueryParamComp];
            if (compQueryParameterValue != null)
            {
                canonicalizedResource.Append(StorageHttpConstants.ConstChars.QuestionMark);
                canonicalizedResource.Append(StorageHttpConstants.QueryParams.QueryParamComp);
                canonicalizedResource.Append(StorageHttpConstants.QueryParams.SeparatorForParameterAndValue);
                canonicalizedResource.Append(compQueryParameterValue);
            }
            
            return canonicalizedResource.ToString();
        }


        /// <summary>
        /// Canonicalize HTTP header contents.
        /// </summary>
        /// <param name="request">An HttpWebRequest object.</param>
        /// <param name="uriComponents">Components of the Uri extracted out of the request.</param>
        /// <returns>The canonicalized string of the given HTTP request's header.</returns>
        internal static string CanonicalizeHttpRequest(HttpWebRequest request, ResourceUriComponents uriComponents)
        {
            return CanonicalizeHttpRequest(
                request.Address, uriComponents, request.Method, request.ContentType, string.Empty, request.Headers);
        }

        /// <summary>
        /// Creates a standard datetime string for the shared key lite authentication scheme.
        /// </summary>
        /// <param name="dateTime">DateTime value to convert to a string in the expected format.</param>
        /// <returns></returns>
        internal static string ConvertDateTimeToHttpString(DateTime dateTime)
        {
            // On the wire everything should be represented in UTC. This assert will catch invalid callers who
            // are violating this rule.
            Debug.Assert(dateTime == DateTime.MaxValue || dateTime == DateTime.MinValue || dateTime.Kind == DateTimeKind.Utc);

            // 'R' means rfc1123 date which is what the storage services use for all dates...
            // It will be in the following format:
            // Sun, 28 Jan 2008 12:11:37 GMT
            return dateTime.ToString("R", CultureInfo.InvariantCulture);
        }

        private static string AppendStringToCanonicalizedString(StringBuilder canonicalizedString, string stringToAppend)
        {
            canonicalizedString.Append(StorageHttpConstants.ConstChars.Linefeed);
            canonicalizedString.Append(stringToAppend);
            return canonicalizedString.ToString();
        }

        internal static string CanonicalizeHttpRequestForSharedKeyLite(HttpWebRequest request, ResourceUriComponents uriComponents, string date)
        {
            StringBuilder canonicalizedString = new StringBuilder(date);
            AppendStringToCanonicalizedString(canonicalizedString, MessageCanonicalizer.GetCanonicalizedResource(request.Address, uriComponents));

            return canonicalizedString.ToString();
        }
    }

    /// <summary>
    /// Use this class to extract various header values from Http requests.
    /// </summary>
    public static class HttpRequestAccessor
    {
        /// <summary>
        /// A helper function for extracting HTTP header values from a NameValueCollection object.
        /// </summary>
        /// <param name="headers">A NameValueCollection object that should contain HTTP header name-values pairs.</param>
        /// <param name="headerName">Name of the header that we want to get values of.</param>
        /// <returns>A array list of values for the header. The values are in the same order as they are stored in the NameValueCollection object.</returns>
        internal static ArrayList GetHeaderValues(NameValueCollection headers, string headerName)
        {
            ArrayList arrayOfValues = new ArrayList();
            string[] values = headers.GetValues(headerName);

            if (values != null)
            {
                foreach (string value in values)
                {
                    // canonization formula requires the string to be left trimmed.
                    arrayOfValues.Add(value.TrimStart());
                }
            }

            return arrayOfValues;
        }


        /// <summary>
        /// Constructs an URI given all its constituents
        /// </summary>
        /// <param name="endpoint">
        /// This is the service endpoint in case of path-style URIs and a host suffix in case of host-style URIs
        /// IMPORTANT: This does NOT include the service name or account name
        /// </param>
        /// <param name="uriComponents">Uri constituents</param>
        /// <param name="pathStyleUri">Indicates whether to construct a path-style Uri (true) or host-style URI (false)</param>
        /// <returns>Full uri</returns>
        public static Uri ConstructResourceUri(Uri endpoint, ResourceUriComponents uriComponents, bool pathStyleUri)
        {
            return pathStyleUri ? 
                    ConstructPathStyleResourceUri(endpoint, uriComponents) : 
                    ConstructHostStyleResourceUri(endpoint, uriComponents);
        }

        /// <summary>
        /// Constructs a path-style resource URI given all its constituents
        /// </summary>
        private static Uri ConstructPathStyleResourceUri(Uri endpoint, ResourceUriComponents uriComponents)
        {
            StringBuilder path = new StringBuilder(string.Empty);
            if (uriComponents.AccountName != null)
            {
                path.Append(uriComponents.AccountName);

                if (uriComponents.ContainerName != null)
                {
                    path.Append(StorageHttpConstants.ConstChars.Slash);
                    path.Append(uriComponents.ContainerName);

                    if (uriComponents.RemainingPart != null)
                    {
                        path.Append(StorageHttpConstants.ConstChars.Slash);
                        path.Append(uriComponents.RemainingPart);
                    }
                }
            }

            return ConstructUriFromUriAndString(endpoint, path.ToString());
        }

        /// <summary>
        /// Constructs a host-style resource URI given all its constituents
        /// </summary>
        private static Uri ConstructHostStyleResourceUri(Uri hostSuffix, ResourceUriComponents uriComponents)
        {
            if (uriComponents.AccountName == null)
            {
                // When there is no account name, full URI is same as hostSuffix
                return hostSuffix;
            }
            else
            {
                // accountUri will be something like "http://accountname.hostSuffix/" and then we append
                // container name and remaining part if they are present.
                Uri accountUri = ConstructHostStyleAccountUri(hostSuffix, uriComponents.AccountName);
                StringBuilder path = new StringBuilder(string.Empty);
                if (uriComponents.ContainerName != null)
                {
                    path.Append(uriComponents.ContainerName);

                    if (uriComponents.RemainingPart != null)
                    {
                        path.Append(StorageHttpConstants.ConstChars.Slash);
                        path.Append(uriComponents.RemainingPart);
                    }
                }
                
                return ConstructUriFromUriAndString(accountUri, path.ToString());
            }
        }


        /// <summary>
        /// Given the host suffix part, service name and account name, this method constructs the account Uri
        /// </summary>
        private static Uri ConstructHostStyleAccountUri(Uri hostSuffix, string accountName)
        {
            // Example: 
            // Input: serviceEndpoint="http://blob.windows.net/", accountName="youraccount"
            // Output: accountUri="http://youraccount.blob.windows.net/"
            Uri serviceUri = hostSuffix;

            // serviceUri in our example would be "http://blob.windows.net/"
            string accountUriString = string.Format(CultureInfo.InvariantCulture,
                                        "{0}{1}{2}.{3}:{4}/",
                                        serviceUri.Scheme,
                                        Uri.SchemeDelimiter,
                                        accountName,
                                        serviceUri.Host,
                                        serviceUri.Port);

            return new Uri(accountUriString);
        }

        private static Uri ConstructUriFromUriAndString(
            Uri endpoint,
            string path)
        {
            // This is where we encode the url path to be valid
            string encodedPath = HttpUtility.UrlPathEncode(path);
            return new Uri(endpoint, encodedPath);
        }
    }

    /// <summary>
    /// Objects of this class contain the credentials (name and key) of a storage account.
    /// </summary>
    public class SharedKeyCredentials
    {

        /// <summary>
        /// Create a SharedKeyCredentials object given an account name and a shared key.
        /// </summary>
        public SharedKeyCredentials(string accountName, byte[] key)
        {
            this.accountName = accountName;
            this.key = key;
        }

        /// <summary>
        /// Signs the request appropriately to make it an authenticated request.
        /// Note that this method takes the URI components as decoding the URI components requires the knowledge
        /// of whether the URI is in path-style or host-style and a host-suffix if it's host-style.
        /// </summary>
        public void SignRequest(HttpWebRequest request, ResourceUriComponents uriComponents)
        {
            if (request == null)
                throw new ArgumentNullException("request");
            string message = MessageCanonicalizer.CanonicalizeHttpRequest(request, uriComponents);
            string computedBase64Signature = ComputeMacSha(message);
            request.Headers.Add(StorageHttpConstants.HeaderNames.Authorization,
                                string.Format(CultureInfo.InvariantCulture,
                                              "{0} {1}:{2}",
                                              StorageHttpConstants.AuthenticationSchemeNames.SharedKeyAuthSchemeName,
                                              accountName,
                                              computedBase64Signature));
        }

        /// <summary>
        /// Signs requests using the SharedKeyLite authentication scheme with is used for the table storage service.
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Lite", 
                                                          Justification = "Name of the authentication scheme in the REST protocol")]        
        public void SignRequestForSharedKeyLite(HttpWebRequest request, ResourceUriComponents uriComponents)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }

            // add the date header to the request
            string dateString = MessageCanonicalizer.ConvertDateTimeToHttpString(DateTime.UtcNow);
            request.Headers.Add(StorageHttpConstants.HeaderNames.StorageDateTime, dateString);

            // compute the signature and add the authentication scheme
            string message = MessageCanonicalizer.CanonicalizeHttpRequestForSharedKeyLite(request, uriComponents, dateString);
            string computedBase64Signature = ComputeMacSha(message);
            request.Headers.Add(StorageHttpConstants.HeaderNames.Authorization,
                                string.Format(CultureInfo.InvariantCulture,
                                              "{0} {1}:{2}",
                                              StorageHttpConstants.AuthenticationSchemeNames.SharedKeyLiteAuthSchemeName,
                                              accountName,
                                              computedBase64Signature));
        }


        private string ComputeMacSha(string canonicalizedString)
        {
            byte[] dataToMAC = System.Text.Encoding.UTF8.GetBytes(canonicalizedString);

            using (HMACSHA256 hmacsha1 = new HMACSHA256(key))
            {
                return System.Convert.ToBase64String(hmacsha1.ComputeHash(dataToMAC));
            }
        }

        private string accountName;
        private byte[] key;
    }
}