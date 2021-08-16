package org.svetikov.reportlaboratorymdf.controller

import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PathVariable
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController

@RestController
@RequestMapping("/api")
class MyController {

    @GetMapping("/{name}")
    fun test(@PathVariable name:String):String{
        return name
    }
}