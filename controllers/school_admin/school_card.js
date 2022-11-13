const SchoolProject = require('../../models/school_admin/SchoolProject')
const SchoolCabinet = require('../../models/school_admin/SchoolCabinet')
const SchoolTeacher = require('../../models/school_admin/SchoolTeacher')
const SchoolCard = require('../../models/school_admin/SchoolCard')
const {validationResult} = require('express-validator')
const error_base = require('../../helpers/error_msg');
const notice_base = require('../../helpers/notice_msg')
const { v4: uuidv4 } = require('uuid');
const FileSaver = require('file-saver');
const xlsx = require('xlsx')
const Blob = require('blob');
const fs = require("fs");
const excel = require('exceljs');
const path = require('path'); 
const http = require('http');
const { Document, Packer, Paragraph, TextRun } = require("docx");
const dateformat = require('../../utils/formatdate')






/**
 *  GET CARD PAGE BY TEACHER ID
 *  school/card/project/2/teacher/id
 *  Раздел с оценками слушателя
 *  */

exports.getCardPageByTeacherId = async (req, res) => {
    try{
        
        if(req.session.user) {

            const school = await SchoolCabinet.getSchoolData(req.session.user)

            const school_name = await school[0].school_name;

            const project = await SchoolProject.getInfoFromProjectById(req.params)

            if(!project.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const projectsIssetSchool = await SchoolProject.getAllProjectsWithThisSchool(req.session.user)
            
            let resultA = []

            for(let i = 0; i < projectsIssetSchool.length; i++) {
                if(project[0].id_project == projectsIssetSchool[i].project_id) {
                    resultA.push(project[0].id_project)
                }
            }

            if(!resultA.length || resultA == 1) {
                return res.status(422).redirect('/school/cabinet');
            } 

            const teacher = await SchoolTeacher.getProfileByTeacherId(req.params)
            if(!teacher.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const disciplineListByTeacherId = await SchoolTeacher.disciplineListByTeacherId(req.params)
            const teacher_id = await teacher[0].id_teacher;

            if(req.params.project_id == 2) {

                if(req.body.id_teacher && req.body._csrf) {

                    let card = await SchoolCard.getCardByTeacherIdWhithFilter(req.body);

                    for(let i = 0; i < card.length; i++) {
                        card[i].sum = card[i].k_1_1_1 + card[i].k_1_1_2 + card[i].k_1_1_3 +card[i].k_1_2_1+card[i].k_2_1_1+card[i].k_2_1_2+card[i].k_2_1_3+card[i].k_2_1_4+card[i].k_2_2_1+card[i].k_2_2_2+card[i].k_2_2_3+card[i].k_2_2_4
                            +card[i].k_2_2_5+card[i].k_2_2_6+card[i].k_2_2_7+card[i].k_2_2_8+card[i].k_2_2_9+card[i].k_2_2_10
                            +card[i].k_2_3_1+card[i].k_2_3_2+card[i].k_2_3_3+card[i].k_2_3_4+
                        card[i].k_2_3_5 + card[i].k_2_3_6 + card[i].k_2_3_7 + card[i].k_2_4_1 + card[i].k_2_4_2 + card[i].k_2_4_3+ card[i].k_2_4_4+ card[i].k_2_4_5
                        + card[i].k_2_4_6+ card[i].k_2_4_7+ card[i].k_2_5_1+ card[i].k_2_5_2+ card[i].k_3_1_1+ card[i].k_3_2_1+ card[i].k_4_1+ card[i].k_4_2+ card[i].k_4_3;
                        card[i].interest = card[i].sum;

     
                         if(card[i].interest > 56 ) {
                             card[i].level = "Выше базового уровня ";
                             card[i].levelStyle = 'success';
                         }else if(card[i].interest < 56 && card[i].interest >= 32) {
                             card[i].level = "Базовый уровень (оптимальный)";
                             card[i].levelStyle = 'good';
                         } else if(card[i].interest < 32) {
                             card[i].level = 'Ниже базового уровня (критический)';
                             card[i].levelStyle = 'danger';
                         }
                     }

                    // for(let i = 0; i < card.length; i++) {
                    //     card[i].sum = card[i].k_1_1 + card[i].k_1_2 + card[i].k_1_3 +
                    //         card[i].k_2_1 + card[i].k_2_2 + card[i].k_3_1 + card[i].k_4_1 + card[i].k_5_1 + card[i].k_5_2 + card[i].k_6_1;
                    //     card[i].interest = card[i].sum * 100 / 20;
                    //
                    //     if(card[i].interest > 84) {
                    //         card[i].level = "Оптимальный уровень";
                    //         card[i].levelStyle = 'success';
                    //     }else if(card[i].interest < 85 && card[i].interest > 59) {
                    //         card[i].level = "Допустимый уровень";
                    //         card[i].levelStyle = 'good';
                    //     }else if(card[i].interest < 60 && card[i].interest > 49) {
                    //         card[i].level = 'критический уровень';
                    //         card[i].levelStyle = 'danger';
                    //     } else if(card[i].interest < 50) {
                    //         card[i].level = 'недопустимый уровень';
                    //         card[i].levelStyle = 'trash';
                    //     }
                    // }

                    const currentSourceId = await card.source;
                    const currentDisc = await card.disc;

                    return res.render('school_teacher_card', {
                        layout: 'maincard',
                        title: 'Личная карта учителя',
                        school_name,
                        teacher,
                        card,
                        teacher_id,
                        school_id: school[0].id_school,
                        disciplineListByTeacherId,
                        project_name: project[0].name_project,
                        project_id: project[0].id_project,
                        currentSourceId,
                        currentDisc,
                        error: req.flash('error'),
                        notice: req.flash('notice')
                    })
                }
             
                let card = await SchoolCard.getCardByTeacherId(req.params);

                for(let i = 0; i < card.length; i++) {
                    card[i].sum = card[i].k_1_1_1 + card[i].k_1_1_2 + card[i].k_1_1_3 +card[i].k_1_2_1+card[i].k_2_1_1+card[i].k_2_1_2+card[i].k_2_1_3+card[i].k_2_1_4+card[i].k_2_2_1+card[i].k_2_2_2+card[i].k_2_2_3+card[i].k_2_2_4
                        +card[i].k_2_2_5+card[i].k_2_2_6+card[i].k_2_2_7+card[i].k_2_2_8+card[i].k_2_2_9+card[i].k_2_2_10
                        +card[i].k_2_3_1+card[i].k_2_3_2+card[i].k_2_3_3+card[i].k_2_3_4+
                        card[i].k_2_3_5 + card[i].k_2_3_6 + card[i].k_2_3_7 + card[i].k_2_4_1 + card[i].k_2_4_2 + card[i].k_2_4_3+ card[i].k_2_4_4+ card[i].k_2_4_5
                        + card[i].k_2_4_6+ card[i].k_2_4_7+ card[i].k_2_5_1+ card[i].k_2_5_2+ card[i].k_3_1_1+ card[i].k_3_2_1+ card[i].k_4_1+ card[i].k_4_2+ card[i].k_4_3;
                    card[i].interest = card[i].sum;


                    if(card[i].interest > 56 ) {
                        card[i].level = "Выше базового уровня ";
                        card[i].levelStyle = 'success';
                    }else if(card[i].interest < 56 && card[i].interest >= 32) {
                        card[i].level = "Базовый уровень (оптимальный)";
                        card[i].levelStyle = 'good';
                    } else if(card[i].interest < 32) {
                        card[i].level = 'Ниже базового уровня (критический)';
                        card[i].levelStyle = 'danger';
                    }
                }

                // for(let i = 0; i < card.length; i++) {
                //    card[i].sum = card[i].k_1_1 + card[i].k_1_2 + card[i].k_1_3 +
                //    card[i].k_2_1 + card[i].k_2_2 + card[i].k_3_1 + card[i].k_4_1 + card[i].k_5_1 + card[i].k_5_2 + card[i].k_6_1;
                //    card[i].interest = card[i].sum * 100 / 20;
                //
                //     if(card[i].interest > 84) {
                //         card[i].level = "Оптимальный уровень";
                //         card[i].levelStyle = 'success';
                //     }else if(card[i].interest < 85 && card[i].interest > 59) {
                //         card[i].level = "Допустимый уровень";
                //         card[i].levelStyle = 'good';
                //     }else if(card[i].interest < 60 && card[i].interest > 49) {
                //         card[i].level = 'критический уровень';
                //         card[i].levelStyle = 'danger';
                //     } else if(card[i].interest < 50) {
                //         card[i].level = 'недопустимый уровень';
                //         card[i].levelStyle = 'trash';
                //     }
                // }


                return res.render('school_teacher_card', {
                    layout: 'maincard',
                    title: 'Личная карта учителя',
                    school_name,
                    teacher,
                    card,
                    teacher_id,
                    school_id: school[0].id_school,
                    disciplineListByTeacherId,
                    project_name: project[0].name_project,
                    project_id: project[0].id_project,
                    error: req.flash('error'),
                    notice: req.flash('notice')
                })
            }else if(req.params.id_project == 3) { 
                console.log('Данный раздел находится в разработке')
                return res.render('admin_page_not_ready', {
                    layout: 'main',
                    title: 'Ошибка',
                    title: 'Предупрехждение',
                    error: req.flash('error'),
                    notice: req.flash('notice')
                })   
            }else {
                console.log('Ошибка в выборе проекта')
                console.log(req.params)
                return res.status(404).render('404_error_template', {
                    layout:'404',
                    title: "Страница не найдена!"});
            }
          }else {
            req.session.isAuthenticated = false
            req.session.destroy( err => {
                if (err) {
                    throw err
                }else {
                    res.redirect('/auth')
                } 
            })
          }
        
    }catch (e) {
        console.log(e)
    }
}

/** END BLOCK */



/** GET ALL CARDS  BY TEACHER ID */

exports.getAllMarksByTeacherId = async (req, res) => {
    try{

        if(req.session.user) {

            const school = await SchoolCabinet.getSchoolData(req.session.user)

            const school_name = await school[0].school_name;

            const project = await SchoolProject.getInfoFromProjectById(req.params)

            if(!project.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const projectsIssetSchool = await SchoolProject.getAllProjectsWithThisSchool(req.session.user)
            
            let resultA = []

            for(let i = 0; i < projectsIssetSchool.length; i++) {
                if(project[0].id_project == projectsIssetSchool[i].project_id) {
                    resultA.push(project[0].id_project)
                }
            }

            if(!resultA.length || resultA == 1) {
                return res.status(422).redirect('/school/cabinet');
            } 

            if(req.params.project_id == 2) {
                console.log(req.params)
                const allMarks = await SchoolCard.getAllMarksByTeacherId(req.params)
                console.log(allMarks)
                if(!allMarks.length) {
                    return res.status(422).redirect('/school/cabinet');
                }
             
                const teacher_id = await allMarks[0].teacher_id;

                const teacher = await SchoolTeacher.getProfileByTeacherId({
                    teacher_id
                })

                const month = ['января', 'февраля','марта', 'апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];

                for(let v = 0; v <allMarks.length; v++ ) {
                    let commonValue = allMarks[v].k_1_1 + allMarks[v].k_1_2 + allMarks[v].k_1_3 + allMarks[v].k_2_1 
                    + allMarks[v].k_2_2 + allMarks[v].k_3_1 + allMarks[v].k_4_1 + allMarks[v].k_5_1 + allMarks[v].k_5_2 
                    + allMarks[v].k_6_1

                    let interest = (commonValue * 100) / 20;

                    allMarks[v].interest  = (commonValue * 100) / 20;
                    

                    let d =  allMarks[v].create_mark_date.getDate();
                    let m =  allMarks[v].create_mark_date.getMonth();

                    let y =  allMarks[v].create_mark_date.getFullYear();
                    allMarks[v].date = `${d}  ${month[m]} ${y}`;
                    allMarks[v].rrrrrrrr = "ccccccccc";
                    let sourceData;

                    if(allMarks[v].source_id == 2) {
                        sourceData = 'Внтуришкольная'
                    }else {
                        sourceData = allMarks[v].source_fio + ', ' + allMarks[v].position_name + ' ( '+ allMarks[v].source_workplace +')'
                    }

                    if(interest > 84) {
                        allMarks[v].level = "Оптимальный уровень";
                        allMarks[v].levelStyle = 'success';
                    }else if(interest < 85 && interest > 59) {
                        allMarks[v].level = "Допустимый уровень";
                        allMarks[v].levelStyle = 'good';
                    }else if(interest < 60 && interest > 49) {
                        allMarks[v].level = 'критический уровень';
                        allMarks[v].levelStyle = 'danger';
                    } else if(interest < 50) {
                        allMarks[v].level = 'недопустимый уровень';
                        allMarks[v].levelStyle = 'trash';
                    }

                    allMarks[v].fio = teacher[0].surname +' '+ teacher[0].firstname + ' ' + teacher[0].patronymic;
                    allMarks[v].position  = allMarks[v].title_position;
                    allMarks[v].school_name = school_name;
                    allMarks[v].commonValue = commonValue;

                }

                const jsonallMarks = JSON.parse(JSON.stringify(allMarks));

                let workbook = new excel.Workbook(); 
                let worksheet = workbook.addWorksheet('allMarks');

                worksheet.columns = [
                    { header: 'ФИО', key: 'fio', width: 10 },
                    { header: 'Должность', key: 'position', width: 30 },
                    { header: 'Предмет', key: 'title_discipline', width: 30},
                    { header: 'Требования Стандартов к предметному содержанию', key: 'k_1_1', width: 50},
                    { header: 'Развитие личностной сферы ученика средствами предмета', key: 'k_1_2', width: 50},
                    { header: 'Использование заданий, развивающих УУД на уроках предмета', key: 'k_1_3', width: 50},
                    { header: 'Учет и развитие мотивации и психофизиологической сферы учащихся', key: 'k_2_1', width: 50},
                    { header: 'Обеспечение целевой психолого-педагогической поддержки обучающихся', key: 'k_2_2', width: 50},
                    { header: 'Требования ЗСС в содержании, структуре урока, в работе с оборудованием и учете данных о детях с ОВЗ', key: 'k_3_1', width: 50},
                    { header: 'Стиль и формы педагогического взаимодействия на уроке', key: 'k_4_1', width: 50},
                    { header: 'Управление организацией учебной деятельности обучающихся через систему оценивания', key: 'k_5_1', width: 50},
                    { header: 'Управление собственной обучающей   деятельностью ', key: 'k_5_2', width: 50},
                    { header: 'Результативность урока', key: 'k_6_1', width: 50},
                    { header: 'Сумма баллов', key: 'commonValue', width: 50},
                    { header: 'Оценка', key: 'level', width: 50},
                    { header: 'Источник ФИО', key: 'source_fio', width: 50},
                    { header: 'Субьект оценивания:', key: 'name_source', width: 30},
                    { header: 'Наименование ОО', key: 'school_name', width: 30}
                ];

                worksheet.addRows(jsonallMarks);

                let excelFileName = teacher[0].surname + '-' + Date.now();

                await workbook.xlsx.writeFile(`files/excels/schools/tmp/${excelFileName}.xlsx`);

                return res.download(path.join(__dirname,'..','..','files','excels','schools','tmp',`${excelFileName}.xlsx`), (err) => {
                   if(err) {
                    console.log('Ошибка при скачивании' + err)
                   }
                })
            }else if(req.params.id_project == 3) {
                console.log('Данный раздел находится в разработке')
                return res.render('admin_page_not_ready', {
                    layout: 'main',
                    title: 'Ошибка',
                    title: 'Предупрехждение',
                    error: req.flash('error'),
                    notice: req.flash('notice')
                })   
            }else {
                console.log('Ошибка в выборе проекта')
                console.log(req.params)
                return res.status(404).render('404_error_template', {
                    layout:'404',
                    title: "Страница не найдена!"});
            }
          }else {
            req.session.isAuthenticated = false
            req.session.destroy( err => {
                if (err) {
                    throw err
                }else {
                    res.redirect('/auth')
                } 
            })
          }
        
    }catch (e) {
        console.log(e)
    }
}

/** END BLOCK */



/** GET CARD PAGE BY TEACHER ID */

exports.addMarkForTeacher = async (req, res) => {
    try{

        if(req.session.user) {
            const project = await SchoolProject.getInfoFromProjectById(req.params)

            if(!project.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const current_date = dateformat();

            const projectsIssetSchool = await SchoolProject.getAllProjectsWithThisSchool(req.session.user)

            let resultA = []

            for(let i = 0; i < projectsIssetSchool.length; i++) {
                if(project[0].id_project == projectsIssetSchool[i].project_id) {
                    resultA.push(project[0].id_project)
                }
            }

            if(!resultA.length || resultA == 1) {
                return res.status(422).redirect('/school/cabinet');
            }

            const teacher = await SchoolTeacher.getProfileByTeacherId(req.params)
            if(!teacher.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const teacher_id = await teacher[0].id_teacher
            const school = await SchoolCabinet.getSchoolData(req.session.user)
            const school_id = await school[0].id_school;
            const school_name = await school[0].school_name;
            const title_area = await school[0].title_area;
            const disciplineListByTeacherId = await SchoolTeacher.disciplineListByTeacherId(req.params)
            const project_id = await project[0].id_project;


            if(req.body.id_teacher && req.body._csrf) {
                const errors = validationResult(req);

                if(!errors.isEmpty()) {
                    req.flash('error', error_base.empty_input)
                    return res.status(422).redirect('/school/card/add/project/' + project_id + '/teacher/' + teacher_id)

                }

                let lastId = await SchoolCard.createNewMarkInCard(req.body);

                console.log('/school/card/project/' + project_id + '/teacher/' + teacher_id)
                    if(lastId) {
                        req.flash('notice', notice_base.success_insert_sql );
                        return res.status(200).redirect('/school/card/project/' + project_id + '/teacher/' + teacher_id);
                    }else {
                        req.flash('error', error_base.wrong_sql_insert)
                        return res.status(422).redirect('/school/card/add/project/' + project_id + '/teacher/' + teacher_id)
                    }
            }

            return res.render('school_teacher_card_add_mark', {
                layout: 'maincard',
                title: 'Оценить урок',
                teacher,
                teacher_id,
                school_id,
                project_id,
                school_name,
                disciplineListByTeacherId,
                title_area,
                current_date,
                error: req.flash('error'),
                notice: req.flash('notice'),
            })
          }else {
            req.session.isAuthenticated = false
            req.session.destroy( err => {
                if (err) {
                    throw err
                }else {
                    res.redirect('/auth')
                }
            })
          }

    }catch (e) {
        console.log(e)
    }
}

/** END BLOCK */



/** GET SINGLE CARD  BY ID */

exports.getSingleCardById = async (req, res) => {
    try{

        if(req.session.user) {

            const school = await SchoolCabinet.getSchoolData(req.session.user)
            const school_id = school[0]['id_school']


            const school_name = await school[0].school_name;

            const project = await SchoolProject.getInfoFromProjectById(req.params)

            if(!project.length) {
                return res.status(422).redirect('/school/cabinet');
            }

            const projectsIssetSchool = await SchoolProject.getAllProjectsWithThisSchool(req.session.user)
            
            let resultA = []

            for(let i = 0; i < projectsIssetSchool.length; i++) {
                if(project[0].id_project == projectsIssetSchool[i].project_id) {
                    resultA.push(project[0].id_project)
                }
            }

            if(!resultA.length || resultA == 1) {
                return res.status(422).redirect('/school/cabinet');
            } 
           
           

            if(req.params.project_id == 2) {

                const singleCard = await SchoolCard.getSingleCard(req.params)

                if(!singleCard.length) {
                    return res.status(422).redirect('/school/cabinet');
                }
             
                const teacher_id = await singleCard[0].teacher_id;


                const teacher = await SchoolTeacher.getProfileByTeacherId({
                    teacher_id
                })



               let commonValue = singleCard[0].k_1_1_1 + singleCard[0].k_1_1_2 + singleCard[0].k_1_1_3 + singleCard[0].k_1_2_1
               + singleCard[0].k_2_1_1 + singleCard[0].k_2_1_2 + singleCard[0].k_2_1_3 + singleCard[0].k_2_1_4 + singleCard[0].k_2_2_1 + singleCard[0].k_2_2_2
               + singleCard[0].k_2_2_3 + singleCard[0].k_2_2_4 + singleCard[0].k_2_2_5 + singleCard[0].k_2_2_6 + singleCard[0].k_2_2_7
               + singleCard[0].k_2_2_8 + singleCard[0].k_2_2_9 + singleCard[0].k_2_2_10 + singleCard[0].k_2_3_1 + singleCard[0].k_2_3_2
               + singleCard[0].k_2_3_3 + singleCard[0].k_2_3_4 + singleCard[0].k_2_3_5 + singleCard[0].k_2_3_6 + singleCard[0].k_2_3_7
               + singleCard[0].k_2_4_1 + singleCard[0].k_2_4_2 + singleCard[0].k_2_4_3 + singleCard[0].k_2_4_4 + singleCard[0].k_2_4_5
               + singleCard[0].k_2_4_6 + singleCard[0].k_2_4_7 + singleCard[0].k_2_5_1 + singleCard[0].k_2_5_2 + singleCard[0].k_3_1_1
               + singleCard[0].k_3_2_1 + singleCard[0].k_4_1 + singleCard[0].k_4_2 + singleCard[0].k_4_3

                let interest = Math.round((commonValue * 100) / 64);

               // let interest = commonValue;

                const d =  singleCard[0].create_mark_date.getDate();
                const m =  singleCard[0].create_mark_date.getMonth();
                const month = ['января', 'февраля','марта', 'апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
                const y =  singleCard[0].create_mark_date.getFullYear();
                const  create_mark_date = `${d}  ${month[m]} ${y}`;
                let sourceData;
                if(singleCard[0].source_id == 2) {
                    sourceData = 'Внтуришкольная'
                }else {
                    sourceData = singleCard[0].source_fio + ', ' + singleCard[0].position_name + ' ( '+ singleCard[0].source_workplace +')'
                }
                let level;
                let levelStyle;
                let levelText;



                if(interest > 75) {
                    level = "Выше базового уровня ";
                    levelStyle = 'success';

                }else if(interest > 50 && interest <= 75) {
                    level = "Базовый уровень (оптимальный)";
                    levelStyle = 'good';

                } else if(interest <= 50) {
                    level = 'Ниже базового уровня (критический)';

                    levelStyle = 'danger';
                }

                singleCard[0].fio = teacher[0].surname +' '+ teacher[0].firstname + ' ' + teacher[0].patronymic;
                singleCard[0].position  = teacher[0].title_position;
                singleCard[0].school_name = school_name;
                singleCard[0].commonValue = commonValue;
                singleCard[0].level = level;


                if(req.body.excel_tbl) {
                    
                    const jsonsingleCard = JSON.parse(JSON.stringify(singleCard));
                    let workbook = new excel.Workbook(); 
                    let worksheet = workbook.addWorksheet('Singlecard');
                    
                    // worksheet.columns = [
                    //     { header: 'ФИО', key: 'fio', width: 10 },
                    //     { header: 'Должность', key: 'position', width: 30 },
                    //     { header: 'Предмет', key: 'title_discipline', width: 30},
                    //     { header: 'Требования Стандартов к предметному содержанию', key: 'k_1_1', width: 50},
                    //     { header: 'Развитие личностной сферы ученика средствами предмета', key: 'k_1_2', width: 50},
                    //     { header: 'Использование заданий, развивающих УУД на уроках предмета', key: 'k_1_3', width: 50},
                    //     { header: 'Учет и развитие мотивации и психофизиологической сферы учащихся', key: 'k_2_1', width: 50},
                    //     { header: 'Обеспечение целевой психолого-педагогической поддержки обучающихся', key: 'k_2_2', width: 50},
                    //     { header: 'Требования ЗСС в содержании, структуре урока, в работе с оборудованием и учете данных о детях с ОВЗ', key: 'k_3_1', width: 50},
                    //     { header: 'Стиль и формы педагогического взаимодействия на уроке', key: 'k_4_1', width: 50},
                    //     { header: 'Управление организацией учебной деятельности обучающихся через систему оценивания', key: 'k_5_1', width: 50},
                    //     { header: 'Управление собственной обучающей   деятельностью ', key: 'k_5_2', width: 50},
                    //     { header: 'Результативность урока', key: 'k_6_1', width: 50},
                    //     { header: 'Сумма баллов', key: 'commonValue', width: 50},
                    //     { header: 'Оценка', key: 'level', width: 50},
                    //     { header: 'Источник ФИО', key: 'source_fio', width: 50},
                    //     { header: 'Субьект оценивания:', key: 'name_source', width: 30},
                    //     { header: 'Наименование ОО', key: 'school_name', width: 30}
                    // ];
                    worksheet.columns = [
                        { header: 'ФИО', key: 'fio', width: 10 },
                        { header: 'Должность', key: 'position', width: 30 },
                        { header: 'Предмет', key: 'title_discipline', width: 30},
                        { header: 'организует изучение (применение) понятия в системе его связей с другими, ранее изученными,понятиями в рамках учебного предмета(раздела, темы)', key: 'k_1_1', width: 50},
                        { header: 'не допускает ошибок, связанных с предметными знаниями и умениями', key: 'k_1_1_2', width: 50},
                        { header: 'указывает на ошибки обучающихся, связанные с предметными знаниями и умениями, оперативно объясняет и исправляет их', key: 'k_1_1_3', width: 50},
                        { header: 'указывает на ошибки обучающихся, связанные с предметными знаниями и умениями,оперативно объясняет и исправляет их', key: 'k_1_2_1', width: 50},
                        { header: 'учитель формулирует цель урока (занятия) вместе с обучающимися, используя проблемную ситуацию (задачу), смысловые догадки, метод ассоциаций и иное', key: 'k_2_1_1', width: 50},
                        { header: 'цель урока (занятия) сформулирована так, что ее достижение можно проверить', key: 'k_2_1_2', width: 50},
                        { header: 'цель урока (занятия) сформулирована четко и понятно для обучающихся', key: 'k_2_1_3', width: 50},
                        { header: 'этапы (задачи) урока (занятия) соответствуют достижению цели, являются необходимыми и достаточными', key: 'k_2_1_4', width: 50},
                        { header: 'учитель использует проблемные методы обучения (частично-поисковый, исследовательский), приемы активизации познавательной деятельности обучающихся, диалоговые технологии', key: 'k_2_2_1', width: 50},
                        { header: 'учитель организует самостоятельную поисковую активность обучающихся', key: 'k_2_2_2', width: 50},
                        { header: 'учитель организует проектную / учебно-исследовательскую деятельность обучающихся или использует проектные и исследовательские задания', key: 'k_2_2_3', width: 50},
                        { header: 'учитель использует задания, которые предусматривают учет индивидуальных особенностей и интересов обучающихся, дифференциацию и индивидуализацию обучения, в том числе возможность выбора темпа, уровня сложности, способов деятельности', key: 'k_2_2_4', width: 50},
                        { header: 'учитель использует задания на формирование / развитие / совершенствование универсальных учебных действий', key: 'k_2_2_5', width: 50},
                        { header: 'учитель использует задания, направленные на формирование положительной учебной мотивации, в том числе учебно-познавательных мотивов', key: 'k_2_2_6', width: 50},
                        { header: 'учитель использует разнообразные способы и средства обратной связи', key: 'k_2_2_7', width: 50},
                        { header: 'качество и количество заданий обеспечивают достижение цели урока (занятия)', key: 'k_2_2_8', width: 50},
                        { header: 'учитель использует методы и приемы, обеспечивающие достижение планируемых результатов урока (занятия)', key: 'k_2_2_9', width: 50},
                        { header: 'выбранный тип урока (занятия) соответствует поставленной цели, структура урока (занятия) логична, этапы взаимосвязаны', key: 'k_2_2_10', width: 50},
                        { header: 'учитель использует формирующее (критериальное) оценивание', key: 'k_2_3_1', width: 50},
                        { header: 'учитель организует разработку / обсуждение критериев оценки деятельности с обучающимися', key: 'k_2_3_2', width: 50},
                        { header: 'учитель организует взаимооценку / самооценку', key: 'k_2_3_3', width: 50},
                        { header: 'учитель содержательно (опираясь на критерии оценки) комментирует выставленные отметки', key: 'k_2_3_4', width: 50},
                        { header: 'учитель организует рефлексию с учетом возрастных особенностей обучающихся (учитель предлагает обучающимся оценить новизну, сложность, полезность выполненных заданий, уровень достижения цели урока (занятия), степень выполнения поставленных задач, полученный результат и деятельность, взаимодействие, иное)', key: 'k_2_3_5', width: 50},
                        { header: 'учитель вместе с детьми оценивает практическую значимость знаний и способов деятельности', key: 'k_2_3_6', width: 50},
                        { header: 'учитель вместе с детьми оценивает практическую значимость знаний и способов деятельности', key: 'k_2_3_7', width: 50},
                        { header: 'учитель использует условно-изобразительную наглядность (знаково-символические средства, модели и др.), использование наглядности целесообразно', key: 'k_2_4_1', width: 50},
                        { header: 'учитель использует ИКТ-технологии, применение технологий целесообразно', key: 'k_2_4_2', width: 50},
                        { header: 'учитель использует наглядность для решения определенной учебной задачи. Средства обучения используются целесообразно с учетом специфики программы, возраста обучающихся', key: 'k_2_4_3', width: 50},
                        { header: 'учитель использует разнообразные справочные материалы (словари, энциклопедии, справочники)', key: 'k_2_4_4', width: 50},
                        { header: 'учитель использует электронные учебные материалы и ресурсы Интернета', key: 'k_2_4_5', width: 50},
                        { header: 'учитель использует материалы разных форматов (текст, таблица, схема, график, видео, аудио)', key: 'k_2_4_6', width: 50},
                        { header: 'учитель использует технологическую карту урока (занятия)', key: 'k_2_4_7', width: 50},
                        { header: 'учитель чередует различные виды деятельности обучающихся', key: 'k_2_5_1', width: 50},
                        { header: 'учитель организует динамические паузы (физкультминутки) и (или) проведение комплекса упражнений для профилактики сколиоза, утомления глаз', key: 'k_2_5_2', width: 50},
                        { header: 'демонстрирует умение решать воспитательные задачи (гражданско-патриотического воспитания, духовно-нравственного воспитания, эстетического воспитания, физического воспитания, формирования культуры здоровья и эмоционального благополучия, трудового воспитания, экологического воспитания,ценности научного познания), используя воспитательные возможности учебного материала урока (занятия)', key: 'k_3_1_1', width: 50},
                        { header: 'учитель использует: задания на морально-этический выбор (моральные дилеммы); задания, организующие рефлексию эмоционального состояния обучающихся; задания, проявляющие ценностное отношение обучающихся', key: 'k_3_2_1', width: 50},
                        { header: 'Учитель демонстрирует умение конструктивно общаться и взаимодействовать с обучающимися на уроке (занятии)', key: 'k_4_1', width: 50},
                        { header: 'Учитель демонстрирует умение управлять общением и взаимодействием обучающихся в совместной деятельности', key: 'k_4_2', width: 50},
                        { header: 'Учитель демонстрирует умение предупреждать и/или решать конфликтные ситуации на уроке (занятии)', key: 'k_4_3', width: 50},
                        { header: 'Сумма баллов', key: 'commonValue', width: 50},
                        { header: 'Оценка', key: 'level', width: 50},
                        { header: 'Источник ФИО', key: 'source_fio', width: 50},
                        { header: 'Субьект оценивания:', key: 'name_source', width: 30},
                        { header: 'Наименование ОО', key: 'school_name', width: 30}
                    ];

                    worksheet.addRows(jsonsingleCard);

                    let dt = new Date();
                    let dtName =  dt.getDate() + '-' + dt.getMonth()+1 +'-'+dt.getFullYear();
                    let excelFileName = teacher[0].surname + '-' + dtName;

                    await workbook.xlsx.writeFile(`files/excels/schools/tmp/${excelFileName}.xlsx`);

                    return res.download(path.join(__dirname,'..','..','files','excels','schools','tmp',`${excelFileName}.xlsx`), (err) => {
                       if(err) {
                        console.log('Ошибка при скачивании' + err)
                       }
                    })
                }


                if(req.body.method_rec) {

                    const jsonsingleCard = JSON.parse(JSON.stringify(singleCard));

                    const k_1_1_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_1_1_1',
                        v_param: jsonsingleCard[0].k_1_1_1
                    });
                    const k_1_1_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_1_1_2',
                        v_param: jsonsingleCard[0].k_1_1_2
                    });
                    
                    const k_1_1_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_1_1_3',
                        v_param: jsonsingleCard[0].k_1_1_3
                    });

                    const k_1_2_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_1_2_1',
                        v_param: jsonsingleCard[0].k_1_2_1
                    });

                    const k_2_1_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_1_1',
                        v_param: jsonsingleCard[0].k_2_1_1
                    });
                    const k_2_1_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_1_2',
                        v_param: jsonsingleCard[0].k_2_1_2
                    });
                    const k_2_1_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_1_3',
                        v_param: jsonsingleCard[0].k_2_1_3
                    });
                    const k_2_1_4_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_1_4',
                        v_param: jsonsingleCard[0].k_2_1_4
                    });

                    const k_2_2_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_1',
                        v_param: jsonsingleCard[0].k_2_2_1
                    });

                    const k_2_2_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_2',
                        v_param: jsonsingleCard[0].k_2_2_2
                    });
                    const k_2_2_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_3',
                        v_param: jsonsingleCard[0].k_2_2_3
                    });

                    const k_2_2_4_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_4',
                        v_param: jsonsingleCard[0].k_2_2_4
                    });
                    const k_2_2_5_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_5',
                        v_param: jsonsingleCard[0].k_2_2_5
                    });
                    const k_2_2_6_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_6',
                        v_param: jsonsingleCard[0].k_2_2_6
                    });
                    const k_2_2_7_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_7',
                        v_param: jsonsingleCard[0].k_2_2_7
                    });
                    const k_2_2_8_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_8',
                        v_param: jsonsingleCard[0].k_2_2_8
                    });
                    const k_2_2_9_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_9',
                        v_param: jsonsingleCard[0].k_2_2_9
                    });
                    const k_2_2_10_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_2_10',
                        v_param: jsonsingleCard[0].k_2_2_10
                    });
                    const k_2_3_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_1',
                        v_param: jsonsingleCard[0].k_2_3_1
                    });
                    const k_2_3_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_2',
                        v_param: jsonsingleCard[0].k_2_3_2
                    });
                    const k_2_3_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_3',
                        v_param: jsonsingleCard[0].k_2_3_3
                    });
                    const k_2_3_4_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_4',
                        v_param: jsonsingleCard[0].k_2_3_4
                    });
                    const k_2_3_5_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_5',
                        v_param: jsonsingleCard[0].k_2_3_5
                    });
                    const k_2_3_6_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_6',
                        v_param: jsonsingleCard[0].k_2_3_6
                    });
                    const k_2_3_7_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_3_7',
                        v_param: jsonsingleCard[0].k_2_3_7
                    });
                    const k_2_4_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_1',
                        v_param: jsonsingleCard[0].k_2_4_1
                    });
                    const k_2_4_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_2',
                        v_param: jsonsingleCard[0].k_2_4_2
                    });
                    const k_2_4_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_3',
                        v_param: jsonsingleCard[0].k_2_4_3
                    });
                    const k_2_4_4_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_4',
                        v_param: jsonsingleCard[0].k_2_4_4
                    });
                    const k_2_4_5_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_5',
                        v_param: jsonsingleCard[0].k_2_4_5
                    });
                    const k_2_4_6_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_6',
                        v_param: jsonsingleCard[0].k_2_4_6
                    });
                    const k_2_4_7_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_4_7',
                        v_param: jsonsingleCard[0].k_2_4_7
                    });
                    const k_2_5_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_5_1',
                        v_param: jsonsingleCard[0].k_2_5_1
                    });
                    const k_2_5_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_2_5_2',
                        v_param: jsonsingleCard[0].k_2_5_2
                    });
                    const k_3_1_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_3_1_1',
                        v_param: jsonsingleCard[0].k_3_1_1
                    });
                    const k_3_2_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_3_2_1',
                        v_param: jsonsingleCard[0].k_3_2_1
                    });

                    const k_4_1_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_4_1',
                        v_param: jsonsingleCard[0].k_4_1
                    });
                    const k_4_2_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_4_2',
                        v_param: jsonsingleCard[0].k_4_2
                    });
                    const k_4_3_rec = await SchoolCard.getRecommendation({
                        k_param: 'k_4_3',
                        v_param: jsonsingleCard[0].k_4_3
                    });



                    const doc = new Document();

                    doc.addSection({
                        children: [
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `ФИО учителя: ${teacher[0].surname} ${teacher[0].firstname} ${teacher[0].patronymic}`,
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                                     }),
                            new Paragraph({
                            children: [
                                new TextRun({
                                    text: `Школа: ${school[0].school_name}`,
                                    bold:true,
                                }),
                                new TextRun('  ')]
                                    }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `город/район: ${school[0].title_area}`,
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun('   '),
                                        ]
                                        }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '                                                    ЗАКЛЮЧЕНИЕ ПО ИТОГУ АНАЛИЗА УРОКА',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun('   '),
                                        ]
                                        }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text:"Категория:",
                                        bold:true
                                    }),
                                    new TextRun({
                                        text: k_1_1_1_rec[0].category,
                                        bold:true
                                    })]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '1.1. Учитель демонстрирует уверенное владение предметным содержанием по теме урока:',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),


                            new Paragraph({
                                children: [
                                            new TextRun('   '),
                                        ]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun({
                                                text: '1.1.1. организует изучение (применение) понятия в системе его связей с другими, ранее изученными, понятиями в рамках учебного предмета (раздела, темы)',
                                                bold:true
                                            })]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun('  '),
                                        ]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun(k_1_1_1_rec[0].content)]
                                        }),
                            new Paragraph({
                                children: [
                                            new TextRun('  '),
                                        ]
                                        }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '1.1.2. не допускает ошибок, связанных с предметными знаниями и умениями',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_1_1_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '1.1.3. указывает на ошибки обучающихся, связанные с предметными знаниями и умениями, оперативно объясняет и исправляет их',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_1_1_3_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '1.2. Учитель демонстрирует умение описывать планируемые предметные результаты урока в деятельностной форме: ',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '1.2.1. описывает предметные учебные действия обучающихся с понятиями, на освоение (или применение) которых направлен урок (занятие)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_1_2_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                //////////////////////////////конец первой группы ПРЕДМЕТНЫЕ

                            new Paragraph({
                                children: [
                                    new TextRun('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text:"Категория: ",
                                        bold:true
                                    }),
                                    new TextRun({
                                        text: k_2_1_1_rec[0].category,
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.1. Учитель демонстрирует умение организовывать целеполагание на уроке (занятии): ',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.1.1. указывает на ошибки обучающихся, связанные с предметными знаниями и умениями, оперативно объясняет и исправляет их',
                                        bold:true
                                    })]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun(k_2_1_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.1.2. цель урока (занятия) сформулирована так, что ее достижение можно проверить',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_1_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.1.3. цель урока (занятия) сформулирована четко и понятно для обучающихся',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_1_3_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.1.4. этапы (задачи) урока (занятия) соответствуют достижению цели, являются необходимыми и достаточными',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_1_4_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2. Учитель демонстрирует умение организовывать разнообразные виды  деятельности обучающихся, которые обеспечивают достижение цели урока (занятия):',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.1. учитель использует проблемные методы обучения (частично-поисковый, исследовательский), приемы активизации познавательной деятельности обучающихся, диалоговые технологии',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.2. учитель организует самостоятельную поисковую активность обучающихся',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.3. учитель организует проектную / учебно-исследовательскую деятельность обучающихся или использует проектные и исследовательские задания',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_3_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.4. учитель использует задания, которые предусматривают учет индивидуальных особенностей и интересов обучающихся, дифференциацию и индивидуализацию обучения, в том числе возможность выбора темпа, уровня сложности, способов деятельности ',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_4_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.5. учитель использует задания на формирование / развитие / совершенствование универсальных учебных действий',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_5_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.6. учитель использует задания, направленные на формирование положительной учебной мотивации, в том числе учебно-познавательных мотивов',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_6_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.7. учитель использует разнообразные способы и средства обратной связи',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_7_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.8. качество и количество заданий обеспечивают достижение цели урока (занятия)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_8_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.9. учитель использует методы и приемы, обеспечивающие достижение планируемых результатов урока (занятия)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_9_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),


                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.2.10. выбранный тип урока (занятия) соответствует поставленной цели, структура урока (занятия) логична, этапы взаимосвязаны',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_2_10_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3. Учитель демонстрирует умение организовывать оценку и рефлексию на уроке (занятии):',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.1. учитель использует формирующее (критериальное) оценивание',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.2. учитель организует разработку / обсуждение критериев оценки деятельности с обучающимися',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.3. учитель организует взаимооценку / самооценку',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_3_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.4. учитель содержательно (опираясь на критерии оценки) комментирует выставленные отметки',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_4_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.5. учитель организует рефлексию с учетом возрастных особенностей обучающихся (учитель предлагает обучающимся оценить новизну, сложность, полезность выполненных заданий, уровень достижения цели урока (занятия), степень выполнения поставленных задач, полученный результат и деятельность, взаимодействие, иное)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_5_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.6. учитель вместе с детьми оценивает практическую значимость знаний и способов деятельности',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_6_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.3.7. соответствие содержания урока (занятия) планируемым результатам',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_3_7_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4. Информационное и техническое обеспечение урока (занятия)',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.1. учитель использует условно-изобразительную наглядность (знаково-символические средства, модели и др.), использование наглядности целесообразно',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.2. учитель использует ИКТ-технологии, применение технологий целесообразно',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),


                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.3. учитель использует наглядность для решения определенной учебной задачи. Средства обучения используются целесообразно с учетом специфики программы, возраста обучающихся',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_3_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.4. учитель использует разнообразные справочные материалы (словари, энциклопедии, справочники)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_4_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),


                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.5. учитель использует электронные учебные материалы и ресурсы Интернета',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_5_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.6. учитель использует материалы разных форматов (текст, таблица, схема, график, видео, аудио)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_6_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.4.7. учитель использует технологическую карту урока (занятия)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_4_7_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.5. Обеспечение условий охраны здоровья обучающихся',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.5.1. учитель чередует различные виды деятельности обучающихся',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_5_1_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '2.5.2. учитель организует динамические паузы (физкультминутки) и (или) проведение комплекса упражнений для профилактики сколиоза, утомления глаз',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_2_5_2_rec[0].content)]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),



                            //////////////////////////////конец второго пункта

                            new Paragraph({
                                children: [
                                    new TextRun('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({

                                children: [
                                    new TextRun({
                                        text:"Категория: ",
                                        bold:true
                                    }),
                                    new TextRun({
                                        text: k_3_1_1_rec[0].category,
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('   '),
                                ]
                            }),


                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '3.1. Учитель объясняет, какие воспитательные задачи достигаются средствами урока (занятия):',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),


                            new Paragraph({
                                children: [
                                            new TextRun({
                                                text: '3.1.1. демонстрирует умение решать воспитательные задачи (гражданско-патриотического воспитания,  духовно-нравственного воспитания, эстетического воспитания, физического воспитания, формирования культуры здоровья и эмоционального благополучия, трудового воспитания, экологического воспитания, ценности научного познания), используя воспитательные возможности учебного материала урока (занятия)',
                                                bold:true
                                            })]
                                        }),
                            new Paragraph({
                            children: [
                                        new TextRun('  '),
                                    ]
                                    }),
                            new Paragraph({
                            children: [
                                        new TextRun(k_3_1_1_rec[0].content)]
                                    }),

                            new Paragraph({
                                children: [
                                            new TextRun('  '),
                                        ]
                                        }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '3.2. Учитель демонстрирует умение организовывать ситуации, развивающие эмоционально-ценностную сферу обучающихся (культуру переживаний и ценностные ориентации)',
                                        bold:true,
                                    }),
                                    new TextRun('  ')]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '3.2.1. учитель использует: задания на морально-этический выбор (моральные дилеммы); задания, организующие рефлексию эмоционального состояния обучающихся; задания, проявляющие ценностное отношение обучающихся ',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_3_2_1_rec[0].content)]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                //////////////////////////////конец третьего пункта

                            new Paragraph({
                                children: [
                                    new TextRun('- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({

                                children: [
                                    new TextRun({
                                        text:"Категория: ",
                                        bold:true
                                    }),
                                    new TextRun({
                                        text: k_4_1_rec[0].category,
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('   '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '4.1. Учитель демонстрирует умение конструктивно общаться и взаимодействовать с обучающимися на уроке (занятии) ',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_4_1_rec[0].content)]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '4.2. Учитель демонстрирует умение управлять общением и взаимодействием обучающихся в совместной деятельности ',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_4_2_rec[0].content)]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: '4.3. Учитель демонстрирует умение предупреждать и/или решать конфликтные ситуации на уроке (занятии)',
                                        bold:true
                                    })]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun(k_4_3_rec[0].content)]
                            }),

                            new Paragraph({
                                children: [
                                    new TextRun('  '),
                                ]
                            }),


        //////////////////////////////конец 10-го пункта

                        ],
                });

                    let excelFileName = teacher[0].surname + '-' + Date.now();

                    await Packer.toBuffer(doc).then((buffer) => {
                         fs.writeFileSync(`files/excels/schools/tmp/${excelFileName}.docx`,buffer,(err) => {
                            if(err) throw err;
                        })
                    });

                    return res.download(path.join(__dirname,'..','..','files','excels','schools','tmp',`${excelFileName}.docx`), (err) => {
                        if(err) {
                         console.log('Ошибка при скачивании' + err)
                        }
                     })
            }

                if(req.body.delete_teacher_school && req.body.delete_teacher_id) {

                    // Удалить оценку в карте учителя

                    req.body['id_card'] = req.params['id_card']
                    req.body['session'] = req.session.user
                    let deleteMarkInCard = await SchoolCard.deleteMarkInCard(req.body);
                    if(deleteMarkInCard.affectedRows) {
                        req.flash('notice', notice_base.success_delete_rows );
                        res.redirect('/school/card/project/'+ req.params['project_id'] + '/teacher/'+req.params['teacher_id']);
                    }else {
                        req.flash('error', error_base.wrong_sql_insert)
                        res.status(422).redirect('/school/card/add/project/' + project_id + '/teacher/' + teacher_id)
                    }
                    // console.log(school[0].id_school)
                }

                return res.render('school_teacher_card_single', {
                    layout: 'maincard',
                    title: 'Личная карта учителя',
                    school_name,
                    teacher,
                    singleCard,
                    teacher_id,
                    commonValue,
                    interest,
                    level,
                    levelStyle,
                    create_mark_date,
                    sourceData,
                    school_id: school[0].id_school,
                    project_name: project[0].name_project,
                    project_id: project[0].id_project,
                    error: req.flash('error'),
                    notice: req.flash('notice')
                })
            }else if(req.params.id_project == 3) {
                console.log('Данный раздел находится в разработке')
                return res.render('admin_page_not_ready', {
                    layout: 'main',
                    title: 'Ошибка',
                    title: 'Предупрехждение',
                    error: req.flash('error'),
                    notice: req.flash('notice')
                })   
            }else {
                console.log('Ошибка в выборе проекта')
                console.log(req.params)
                return res.status(404).render('404_error_template', {
                    layout:'404',
                    title: "Страница не найдена!"});
            }
          }else {
            req.session.isAuthenticated = false
            req.session.destroy( err => {
                if (err) {
                    throw err
                }else {
                    res.redirect('/auth')
                } 
            })
          }
        
    }catch (e) {
        console.log(e)
    }
}

/** END BLOCK */







