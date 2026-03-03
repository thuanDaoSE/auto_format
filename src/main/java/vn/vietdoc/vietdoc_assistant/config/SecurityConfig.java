package vn.vietdoc.vietdoc_assistant.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.web.SecurityFilterChain;

@Configuration
public class SecurityConfig {

    @Bean
    public SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
        http
            .csrf(csrf -> csrf.disable()) // Tắt chống giả mạo request (cần thiết cho API)
            .authorizeHttpRequests(auth -> auth
                .requestMatchers("/api/format/**").permitAll() // Mở cửa tự do cho API định dạng Word
                .anyRequest().permitAll() // Tạm thời mở cửa hết cho mọi request
            );
        return http.build();
    }
}