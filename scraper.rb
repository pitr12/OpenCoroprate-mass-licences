# -*- coding: utf-8 -*-

require 'json'
require 'turbotlib'
require 'nokogiri'
require 'faraday'
require 'roo'

def self.get_xls_links(cell)
  body = (Faraday.get 'http://www.mass.gov/ocabr/licensee/license-types/insurance/individual-and-business-entity-licensing/licensed-listings.html').body
  site = Nokogiri::HTML(body)
  list = []
  table = site.search('table')[1]
  table.search('tr').each do |tr|
    td = tr.search('td')[cell]
    l = td.search('a').map {|link| "http://www.mass.gov" + link['href']}
    list << l[2] if !l.empty?
  end
  list
end

Turbotlib.log("Starting run...") # optional debug logging

#cell 1 = agencies, cell 2 = individuals
XLS_AGENCIES_LINKS = get_xls_links(1)
XLS_INDIVIDUALS_LINKS = get_xls_links(2)

XLS_AGENCIES_LINKS.each do |xls_link|
  xls = Roo::Excel.new(xls_link)
    (xls.first_row..xls.last_row).each do |row_number|
      row = xls.row(row_number)
      if row[0] != nil
        data = {
            number: row[0].to_i,
            license: row[1],
            licensure: row[4],
            agency: row[6],
            adress: row[8],
            city: row[9],
            state: row[11],
            zip: row[12],
            phone: row[16],
            source_url: xls_link
        }
       #puts JSON.dump(data)
      end
    end
end

XLS_INDIVIDUALS_LINKS.each do |xls_link|
  xls = Roo::Excel.new(xls_link)
  (xls.first_row..xls.last_row).each do |row_number|
    row = xls.row(row_number)
    if row[0] != nil
      data = {
          number: row[0].to_i,
          license: row[1],
          licensure: row[3],
          individual: row[5],
          adress: row[7],
          city: row[8],
          state: row[9],
          zip: row[10],
          phone: row[12],
          email: row[14],
          source_url: xls_link
      }
      puts JSON.dump(data)
    end
  end
end
