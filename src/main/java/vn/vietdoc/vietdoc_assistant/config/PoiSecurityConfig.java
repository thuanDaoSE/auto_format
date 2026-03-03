package vn.vietdoc.vietdoc_assistant.config;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.context.annotation.Configuration;

import jakarta.annotation.PostConstruct;

@Configuration
public class PoiSecurityConfig {
    @PostConstruct
    public void init() {
        ZipSecureFile.setMinInflateRatio(0.001); 
        ZipSecureFile.setMaxEntrySize(50L * 1024L * 1024L);
        ZipSecureFile.setMaxTextSize(15L * 1024L * 1024L);
    }
}