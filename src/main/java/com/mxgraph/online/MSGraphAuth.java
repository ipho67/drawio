/**
 * Copyright (c) 2006-2025, JGraph Holdings Ltd, draw.io AG
 */
package com.mxgraph.online;

@SuppressWarnings("serial")
abstract public class MSGraphAuth extends AbsAuth
{
	public static String CLIENT_SECRET_FILE_PATH = "msgraph_client_secret";
	public static String CLIENT_ID_FILE_PATH = "msgraph_client_id";
	public static String TENANT_ID_FILE_PATH = "msgraph_tenant_id";

	private static volatile Config CONFIG = null;

	protected Config getConfig()
	{
		if (CONFIG == null)
		{
			synchronized (MSGraphAuth.class)
			{
				if (CONFIG == null)
				{
					String clientSecrets = SecretFacade.getSecret(CLIENT_SECRET_FILE_PATH, getServletContext());
					String clientId = SecretFacade.getSecret(CLIENT_ID_FILE_PATH, getServletContext());
					String tenantId = null;

					try
					{
						tenantId = SecretFacade.getSecret(TENANT_ID_FILE_PATH, getServletContext());
					}
					catch (Exception e)
					{
						// The tenant setting is optional; use the Microsoft common endpoint.
					}

					String tenantIdPathPart = (tenantId != null && !tenantId.trim().isEmpty())
							? tenantId.trim() : "common";

					// A tenant identifier is a single URL path segment, never a URL or path.
					if (!tenantIdPathPart.matches("[A-Za-z0-9](?:[A-Za-z0-9.-]{0,251}[A-Za-z0-9])?"))
					{
						throw new IllegalStateException("Invalid Microsoft tenant identifier");
					}

					Config config = new Config(clientId, clientSecrets);
					config.AUTH_SERVICE_URL = "https://login.microsoftonline.com/" + tenantIdPathPart + "/oauth2/v2.0/token";
					config.REDIRECT_PATH = "/microsoft";

					// Publish only after the configuration is fully initialized.
					CONFIG = config;
				}
			}
		}

		return CONFIG;
	}

	public MSGraphAuth()
	{
		super();
		cookiePath = "/microsoft";
	}

	protected String processAuthResponse(String authRes, boolean jsonResponse)
	{
		StringBuilder res = new StringBuilder();

		if (!jsonResponse)
		{
			res.append("<!DOCTYPE html><html><head><script>");
			res.append("(function() { var authInfo = ");
		}

		// authRes is a JSON object constructed server-side from the Microsoft token response.
		res.append(authRes);

		if (!jsonResponse)
		{
			res.append(";");
			res.append("var trustedOpener = null;");
			res.append("try {");
			res.append(" if (window.opener != null &&");
			res.append(" window.opener.location.origin === window.location.origin &&");
			res.append(" typeof window.opener.onOneDriveCallback === 'function') {");
			res.append(" trustedOpener = window.opener;");
			res.append(" }");
			res.append("} catch (e) { /* Cross-origin opener: do not expose auth data. */ }");
			res.append("if (trustedOpener != null) {");
			res.append(" trustedOpener.onOneDriveCallback(authInfo, window);");
			res.append("} else {");
			res.append(" var head = document.getElementsByTagName('head')[0];");
			res.append(" var script = document.createElement('script');");
			res.append(" script.onload = function() {");
			res.append("  var authInfoStr = JSON.stringify(authInfo);");
			res.append("  try { localStorage.setItem('.oneDriveAuthInfo', '{}'); } catch (e) {}");
			res.append("  Office.onReady(function () {");
			res.append("   var origin = window.location.origin;");
			res.append("   if ((window.location.protocol === 'https:' || window.location.protocol === 'http:') &&");
			res.append("       origin !== 'null' && Office.context != null && Office.context.ui != null) {");
			res.append("    Office.context.ui.messageParent(authInfoStr, { targetOrigin: origin });");
			res.append("   }");
			res.append("  });");
			res.append(" };");
			res.append(" script.src = 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js';");
			res.append(" head.appendChild(script);");
			res.append("}");
			res.append("})();</script></head><body><div>Automatic login interrupted. Please close and select OneDrive again.</div></body></html>");
		}

		return res.toString();
	}
}
