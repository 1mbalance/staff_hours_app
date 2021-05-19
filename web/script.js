function check_db() {
    eel.check_db()
}

function no_db(text) {
    var div = document.getElementById("no_db");
    div.innerHTML = text;
}
eel.expose(no_db);

function db(text) {
    var text_area = document.getElementById("db");
    text_area.value += text;
}
eel.expose(db);

function fill_db() {
    eel.fill_db()
}

function update_db() {
    eel.update_db()
}

function get_time() {
    eel.get_time()
}

function delete_db() {
    eel.delete_db()
}

function change_rank() {
    var ranks = new Array();
    ranks[0] = 'Помощник';
    ranks[1] = 'Модератор';
    ranks[2] = 'Администратор';
    ranks[3] = 'Старший администратор';
    ranks[4] = 'Куратор';
    ranks[5] = 'Cупер администратор';

    var options = '';

    for (var i = 0; i < ranks.length; i++) {
      options += '<option value="' + ranks[i] + '" />';
    }

    document.getElementById('change_rank').innerHTML = options;
}

function change_admin() {
    var number_input = document.getElementById('choose_admin')
    admin_id = number_input.value;
    var text_input = document.getElementById('change_nick')
    admin_new_nick = text_input.value;
    var datalist_input = document.getElementById('change_rank_input')
    admin_new_rank = datalist_input.value;

    eel.change_admin(admin_id, admin_new_nick, admin_new_rank)
}

function delete_admin() {
    var number_input = document.getElementById('choose_admin')
    admin_id = number_input.value;

    eel.delete_admin(admin_id)
}

function check_admin() {
    eel.check_admin()
}

function check_admin_time() {
    eel.check_admin_time()
}

function no_admin_time(text) {
    var div = document.getElementById("no_admin_time");
    div.innerHTML = text;
}
eel.expose(no_admin_time);

function check_time() {
    var number_input = document.getElementById('choose_admin_2')
    admin_id = number_input.value;

    eel.check_time(admin_id)
}

function check_time_2() {
    eel.check_time_2()
}

function admin_time(date) {
    var text_area = document.getElementById("admin_time");
    text_area.value += date;
}
eel.expose(admin_time);

function admin_week_time() {
    var text_input = document.getElementById('date_1')
    date_1 = text_input.value;
    var text_input_2 = document.getElementById('date_2')
    date_2 = text_input_2.value;

    eel.admin_week_time(date_1, date_2)
}

function admin_week_time_2() {
    eel.admin_week_time_2()
}

function admin_time_res(text, date_1, date_2) {
    var div = document.getElementById("admin_time_res");
    div.innerHTML = 'Результат(c ' + date_1 + ' по '+ date_2 + '): ' + text;
}
eel.expose(admin_time_res);

function check_vacation() {
    eel.check_vacation()
}

function get_vacation(start_vacation, end_vacation) {
    var div = document.getElementById("present_vacation");
    div.innerHTML = 'Имеется отпуск c ' + start_vacation + ' по '+ end_vacation + '.';
}
eel.expose(get_vacation);

function give_vacation() {
    var date_input = document.getElementById('start_vacation')
    start_vacation = date_input.value;
    var date_input_2 = document.getElementById('end_vacation')
    end_vacation = date_input_2.value;

    eel.give_vacation(start_vacation, end_vacation)
}

function delete_vacation() {
    eel.delete_vacation()
}

window.onload = async function date_list () {
  let date_list = await eel.date_list()();

  var options_2 = '';

    for (var i = 0; i < date_list.length; i++) {
      options_2 += '<option value="' + date_list[i] + '" />';
    }

    document.getElementById('date_list').innerHTML = options_2;
    document.getElementById('date_list_2').innerHTML = options_2;

}

function create_excel() {
    eel.create_excel_prep()
}

function create_excel_2() {
    var list_input = document.getElementById('date_list_input')
    date_1 = list_input.value;
    var list_input_2 = document.getElementById('date_list_input_2')
    date_2 = list_input_2.value;

    eel.create_excel(date_1, date_2)
}