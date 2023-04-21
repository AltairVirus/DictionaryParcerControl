import Config
import psycopg2


def db_connect(data_line='ri.str_value', db='genres', look_for=True):
    try:
        if db == 'genres':
            connection = psycopg2.connect(user=Config.gr_user,
                                          password=Config.gr_password,
                                          host=Config.gr_host,
                                          port=Config.gr_port,
                                          database=Config.gr_database)

            cursor = connection.cursor()
            print(f'Соединение с PostgreSQL -- {db} -- открыто')

            if look_for:
                postgreSQL_select_Keys = f'''
                SELECT
                json_agg(DISTINCT {data_line})
                from model.rule_group rg
                JOIN model.objects o ON rg.global_id = o.global_id
                JOIN model.rules_where rw ON rg.global_id = rw.rule_group_id
                LEFT JOIN model.rule_item ri ON rw.global_id = ri.rules_where_id
                LEFT JOIN model.rule_values rv ON rw.global_id = rv.rules_where_id
                WHERE o.global_id = '{Config.global_id}'
                GROUP by o.name_short_ru, ri.rules_where_id'''

            else:
                postgreSQL_select_Keys = f'''
                select ri.rules_where_id as "id строки",
                ri.rule_attr as "id атрибута",
                ri.str_value as "значение атрибута"
                from model.objects o
                JOIN model.rules_where rw ON o.global_id = rw.rule_group_id
                LEFT JOIN model.rule_item ri ON rw.global_id = ri.rules_where_id
                where o.name = '{Config.object_name}' 
                and ri.rule_attr is not null'''

        elif db == 'opm':
            connection = psycopg2.connect(user=Config.opm_user,
                                          password=Config.opm_password,
                                          host=Config.opm_host,
                                          port=Config.opm_port,
                                          database=Config.opm_database)

            cursor = connection.cursor()
            print(f'Соединение с PostgreSQL -- {db} -- открыто')

            postgreSQL_select_Keys = f'''
            SELECT p.id AS "id атрибута",
            p.name AS "имя атрибута"
            FROM public.property p'''

        cursor.execute(postgreSQL_select_Keys)

        print(f"Данные {data_line} переданы")
        return cursor.fetchall()


    except(Exception, psycopg2.Error) as error:
        print("Ошибка соединения с PostgreSQL", error)

    finally:
        if connection:
            cursor.close()
            connection.close()
            print("Соединение с PostgreSQL закрыто")
