package com.github.gn5r.poi.utils;

import com.github.gn5r.poi.util.annotation.Cell;

import lombok.Data;

public class PoiUtilTest {

    @Data
    public class User {
        @Cell(tags = {"name"})
        private String name;

        @Cell(tags = {"age"})
        private Integer age;

        @Cell(tags = {"sex"})
        private String sex;

        @Cell(tags = {"birthday"})
        private String birthday;

        public void setTestUser() {
            this.name = "テスト";
            this.age = 24;
            this.sex = "男";
            this.birthday = "1996/11/02";
        }
    }
}
