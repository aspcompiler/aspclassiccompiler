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
// <copyright file="Configuration.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Configuration;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using Microsoft.Samples.ServiceHosting.StorageClient;
using Microsoft.ServiceHosting.ServiceRuntime;

[assembly: CLSCompliant(true)]

namespace Microsoft.Samples.ServiceHosting.AspProviders
{
    internal static class Configuration
    {

        internal const string DefaultMembershipTableNameConfigurationString = "DefaultMembershipTableName";
        internal const string DefaultRoleTableNameConfigurationString = "DefaultRoleTableName";
        internal const string DefaultSessionTableNameConfigurationString = "DefaultSessionTableName";
        internal const string DefaultSessionContainerNameConfigurationString = "DefaultSessionContainerName";
        internal const string DefaultProfileContainerNameConfigurationString = "DefaultProfileContainerName";
        internal const string DefaultProviderApplicationNameConfigurationString = "DefaultProviderApplicationName";

        internal const string DefaultMembershipTableName = "Membership";
        internal const string DefaultRoleTableName = "Roles";
        internal const string DefaultSessionTableName = "Sessions";
        internal const string DefaultSessionContainerName = "sessionprovidercontainer";
        internal const string DefaultProfileContainerName = "profileprovidercontainer";
        internal const string DefaultProviderApplicationName = "appname";


        internal static string GetConfigurationSetting(string configurationString, string defaultValue)
        {
            return GetConfigurationSetting(configurationString, defaultValue, false);
        }

        /// <summary>
        /// Gets a configuration setting from application settings in the Web.config or App.config file. 
        /// When running in a hosted environment, configuration settings are read from the settings specified in 
        /// .cscfg files (i.e., the settings are read from the fabrics configuration system).
        /// </summary>
        internal static string GetConfigurationSetting(string configurationString, string defaultValue, bool throwIfNull)
        {
            if (string.IsNullOrEmpty(configurationString)) {
                throw new ArgumentException("The parameter configurationString cannot be null or empty.");
            }

            string ret = null;

            // first, try to read from appsettings
            ret = TryGetAppSetting(configurationString);

            // settings in the csc file overload settings in Web.config
            if (RoleManager.IsRoleManagerRunning)
            {
                string cscRet = TryGetConfigurationSetting(configurationString);
                if (!string.IsNullOrEmpty(cscRet))
                {
                    ret = cscRet;
                }

                // if there is a csc config name in the app settings, this config name even overloads the 
                // setting we have right now
                string refWebRet = TryGetAppSetting(StorageAccountInfo.CSConfigStringPrefix + configurationString);
                if (!string.IsNullOrEmpty(refWebRet))
                {
                    cscRet = TryGetConfigurationSetting(refWebRet);
                    if (!string.IsNullOrEmpty(cscRet))
                    {
                        ret = cscRet;
                    }
                }
            }

            // if we could not retrieve any configuration string set return value to the default value
            if (string.IsNullOrEmpty(ret) && defaultValue != null)
            {
                ret = defaultValue;
            }

            if (string.IsNullOrEmpty(ret) && throwIfNull)
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "Cannot find configuration string {0}.", configurationString));
            }
            return ret;
        }

        internal static string GetConfigurationSettingFromNameValueCollection(NameValueCollection config, string valueName)
        {
            if (config == null)
            {
                throw new ArgumentNullException("config");
            }
            if (valueName == null) {
                throw new ArgumentNullException("valueName");
            }

            string sValue = config[valueName];

            if (RoleManager.IsRoleManagerRunning)
            {
                // settings in the hosting configuration are stronger than settings in app config
                string cscRet = TryGetConfigurationSetting(valueName);
                if (!string.IsNullOrEmpty(cscRet))
                {
                    sValue = cscRet;
                }

                // if there is a csc config name in the app settings, this config name even overloads the 
                // setting we have right now
                string refWebRet = config[StorageAccountInfo.CSConfigStringPrefix + valueName];
                if (!string.IsNullOrEmpty(refWebRet))
                {
                    cscRet = TryGetConfigurationSetting(refWebRet);
                    if (!string.IsNullOrEmpty(cscRet))
                    {
                        sValue = cscRet;
                    }
                }
            }
            return sValue;
        }

        internal static bool GetBooleanValue(NameValueCollection config, string valueName, bool defaultValue)
        {
            string sValue = GetConfigurationSettingFromNameValueCollection(config, valueName);

            if (string.IsNullOrEmpty(sValue))
            {
                return defaultValue;
            }

            bool result;
            if (bool.TryParse(sValue, out result))
            {
                return result;
            }
            else
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value must be boolean (true or false) for property '{0}'.", valueName));
            }
        }

        internal static int GetIntValue(NameValueCollection config, string valueName, int defaultValue, bool zeroAllowed, int maxValueAllowed)
        {
            string sValue = GetConfigurationSettingFromNameValueCollection(config, valueName);

            if (string.IsNullOrEmpty(sValue))
            {
                return defaultValue;
            }

            int iValue;
            if (!Int32.TryParse(sValue, out iValue))
            {
                if (zeroAllowed)
                {
                    throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value must be a non-negative 32-bit integer for property '{0}'.", valueName));
                }

                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value must be a positive 32-bit integer for property '{0}'.", valueName));
            }

            if (zeroAllowed && iValue < 0)
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value must be a non-negative 32-bit integer for property '{0}'.", valueName));
            }

            if (!zeroAllowed && iValue <= 0)
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value must be a positive 32-bit integer for property '{0}'.", valueName));
            }

            if (maxValueAllowed > 0 && iValue > maxValueAllowed)
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The value '{0}' can not be greater than '{1}'.", valueName, maxValueAllowed.ToString(CultureInfo.InstalledUICulture)));
            }

            return iValue;
        }

        internal static string GetStringValue(NameValueCollection config, string valueName, string defaultValue, bool nullAllowed)
        {
            string sValue = GetConfigurationSettingFromNameValueCollection(config, valueName); 

            if (string.IsNullOrEmpty(sValue) && nullAllowed)
            {
                return null;
            }
            else if (string.IsNullOrEmpty(sValue) && defaultValue != null)
            {
                return defaultValue;
            }
            else if (string.IsNullOrEmpty(sValue))
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The parameter '{0}' must not be empty.", valueName));
            }

            return sValue;
        }


        internal static string GetStringValueWithGlobalDefault(NameValueCollection config, string valueName, string defaultConfigString, string defaultValue, bool nullAllowed)
        {
            string sValue = GetConfigurationSettingFromNameValueCollection(config, valueName);

            if (string.IsNullOrEmpty(sValue))
            {
                sValue = GetConfigurationSetting(defaultConfigString, null);
            }

            if (string.IsNullOrEmpty(sValue) && nullAllowed)
            {
                return null;
            }
            else if (string.IsNullOrEmpty(sValue) && defaultValue != null)
            {
                return defaultValue;
            }
            else if (string.IsNullOrEmpty(sValue))
            {
                throw new ConfigurationErrorsException(string.Format(CultureInfo.InstalledUICulture, "The parameter '{0}' must not be empty.", valueName));
            }

            return sValue;
        }

        internal static string GetInitExceptionDescription(StorageAccountInfo table, StorageAccountInfo blob) {
            StringBuilder builder = new StringBuilder();
            builder.Append(GetInitExceptionDescription(table, "table storage configuration"));
            builder.Append(GetInitExceptionDescription(blob, "blob storage configuration"));
            return builder.ToString();
        }

        internal static string GetInitExceptionDescription(StorageAccountInfo info, string desc) {
            StringBuilder builder = new StringBuilder();
            builder.Append("The reason for this exception is typically that the endpoints are not correctly configured. " + Environment.NewLine);
            if (info == null)
            {
                builder.Append("The current " + desc + " is null. Please specify a table endpoint!" + Environment.NewLine);
            }
            else
            {
                string baseUri = (info.BaseUri == null) ? "null" : info.BaseUri.ToString();                
                builder.Append("The current " + desc + " is: " + baseUri + ", usePathStyleUris is: " + info.UsePathStyleUris + Environment.NewLine);
                builder.Append("Please also make sure that the account name and the shared key are specified correctly. This information cannot be shown here because of security reasons.");
            }
            return builder.ToString();
        }

        private static string TryGetConfigurationSetting(string configName)
        {
            string ret = null;
            try
            {
                ret = RoleManager.GetConfigurationSetting(configName);
            }
            catch (RoleException)
            {
                return null;
            }
            return ret;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage ("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", 
            Justification = "Make sure that no error condition prevents environment from reading service configuration.")]
        private static string TryGetAppSetting(string configName)
        {
            string ret = null;
            try
            {
                ret = ConfigurationSettings.AppSettings[configName];
            }
            // some exception happened when accessing the app settings section
            // most likely this is because there is no app setting file
            // this is not an error because configuration settings can also be located in the cscfg file, and explicitly 
            // all exceptions are captured here
            catch (Exception)
            {
                return null;
            }
            return ret;
        }

    }
}