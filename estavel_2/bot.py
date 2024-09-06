def main():
    # Preparar a pasta export_data
    prepare_export_folder("export_data")

    attachments = []
    for i in range(len(orgs)):
        latest_file = None
        gerp_web(webbot, orgs[i], app_name)
        helpers.keep_download_gerp_executable(webbot, i+1)
        executable_gerp = webbot.get_last_created_file()
        
        if executable_gerp[-15:] != "frmservlet.jnlp": return
        desktop_bot.execute(executable_gerp)
        
        helpers.pass_warning_open_java_app(desktop_bot)
        helpers.activate_gerp_window(desktop_bot)
    
        att = dowload_files_from_gerp(desktop_bot, orgs[i], latest_file)
        print("ARQUIVOS => ", att)
        attachments.append(att)
        
        desktop_bot.wait(4000)
        os.system("taskkill /f /im jp2launcher.exe")

    # Zipar os arquivos da pasta export_data
    zip_file = zip_files("export_data")
    attachments.append(zip_file)
    
    is_test = '1'    
    common.send_email_is_test(is_test, attachments)
