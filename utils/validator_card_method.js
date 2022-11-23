const {body} = require('express-validator')

exports.validatorForAddMarkMethod  = [
    body('discipline_id','Заполните все обязательные поля!').toInt().isInt(),
    body('class_id','Заполните все обязательные поля!').notEmpty().toInt().isInt(),
    body('liter_class').escape().trim(),
    body('source_id','Заполните все обязательные поля!').toInt().isInt(),
    body('id_teacher','Заполните все обязательные поля!').notEmpty().trim(),
    body('school_id','Заполните все обязательные поля!').notEmpty().toInt().isInt(),
    body('project_id','Заполните все обязательные поля!').notEmpty().toInt().isInt(),
    body('thema','Заполните все обязательные поля!').notEmpty().escape().trim(),
    body('k_2_1_1','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_1_2','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_1_3','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_1_4','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_1','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_2','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_3','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_4','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_5','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_6','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_7','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_8','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_9','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_2_10','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_1','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_2','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_3','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_4','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_5','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_6','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_3_7','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_1','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_2','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_3','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_4','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_5','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_6','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_4_7','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_5_1','Заполните все обязательные поля!').toInt().isInt(),
    body('k_2_5_2','Заполните все обязательные поля!').toInt().isInt(),


    // body('total_experience').custom( (value, {req}) => {
    //     if(parseInt(value) < parseInt(req.body.teaching_experience)) {
    //         throw new Error('Общий стаж не может быть меньше педагогического')
    //     }
    // })

]